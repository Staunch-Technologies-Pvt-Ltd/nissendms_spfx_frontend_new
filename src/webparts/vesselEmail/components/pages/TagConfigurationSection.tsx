import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import { dmsBtn, dmsControlStyle } from '../dmsDesignSystem';

/**
 * Settings → Tag Configuration.
 * Backend: backend/app/tag_config_api.py (+ services/tag_config.py), docs/tag-configuration.md.
 * Levels: Domain > Main Folder > Group > Category > Sub Category (levels may be skipped).
 * Vessel Names are read-only (Term Store).
 */

type Level = 'domain' | 'main_folder' | 'group' | 'category' | 'sub_category';
type Tab = Level | 'vessel_names' | 'internal';
type Mode = 'add' | 'replace';
// 'site' = the session's current active site, 'template' = the default for
// new clients, or a literal site_key picked from the site selector below —
// see GET /api/tag-config/sites.
type Scope = string;

interface TagItem {
  id: number; level: Level; name: string; display_name: string; folder_name: string;
  code: string | null; description: string | null; parent_id: number | null; sort_order: number;
  status: 'Active' | 'Inactive' | 'Archived'; is_default: boolean; source: string; path: string;
  aliases: string[];
}
interface ConfigResponse {
  site_key: string; origin: string; scope: Scope; items: TagItem[];
  levels: Array<{ id: Level; label: string }>; modes: Record<Level, Mode>; is_admin: boolean;
  folder_name_rules: { invalid_chars: string; max_length: number; ampersand: string };
}
interface DiffEntry { action: 'new' | 'unchanged' | 'reactivate' | 'deactivate' | 'error'; level: string; name: string; parent_path: string; message: string; in_use: number; line: number; }
interface DiffResult { mode: Mode; counts: Record<string, number>; confirmation: string; entries: DiffEntry[]; }
interface CopyDiffResult extends DiffResult { source_label?: string; target_label?: string; }
interface Snapshot { id: number; created_at: string | null; changed_by: string | null; mode: string; level: string | null; summary: Record<string, unknown>; }
interface SyncResult { level?: string; column: { displayName?: string } | null; type: string | null; current: string[]; target: string[]; add: string[]; remove: string[]; applied: boolean; error: string | null;
  /** Domain level only: Domains that are (or would be) created as term sets in the Vessel DMS Term Store group. */
  term_store?: { add: string[]; added: string[]; error: string | null }; }
interface RetagResult { level?: string; old_value: string; new_value: string | null; checked: number; updated: number; errors: string[]; }
const SYNCABLE_LEVELS: Level[] = ['domain', 'group', 'category'];
interface ImportRow { Level: string; 'Parent Path': string; Name: string; Code?: string; Description?: string; 'Sort Order'?: number | null; Status?: string; }
interface SiteOption { site_key: string; display_name: string; }

const LEVELS: Level[] = ['domain', 'main_folder', 'group', 'category', 'sub_category'];
const LABEL: Record<Tab, string> = {
  domain: 'Domains', main_folder: 'Main Folders', group: 'Groups', category: 'Categories',
  sub_category: 'Sub Categories', vessel_names: 'Vessel Names', internal: 'Internal',
};
const SINGULAR: Record<Level, string> = {
  domain: 'Domain', main_folder: 'Main Folder', group: 'Group', category: 'Category', sub_category: 'Sub Category',
};
const INVALID = /[~"#%*:<>?/\\{|}]/;

// Simplified admin view. Set to false to bring back the advanced controls
// (client scope picker, Add/Replace mode, Source filter, List/Tree toggle,
// Code/Description fields, Copy-from-site, Reset to Default).
const SIMPLE_UI = true;
// Levels that only drive folder-path parsing and have no SharePoint column,
// so they are shown read-only in simple view.
const INTERNAL_LEVELS: Level[] = ['main_folder', 'sub_category'];
const SIMPLE_TABS: Tab[] = ['domain', 'group', 'category', 'vessel_names', 'internal'];

// Token-backed palette — every value resolves through `clay.*` / `var(--vdms-*)`
// so this section matches Dashboard's look (and re-themes with it).
const C = {
  text: 'var(--vdms-text)', sub: 'var(--vdms-text-muted)', border: 'var(--vdms-line)', blue: clay.accent, blueBg: clay.accentSoft,
  amberBg: clay.pillWarnBg, amber: clay.pillWarnText, redBg: clay.pillDangerBg, red: clay.pillDangerText, greenBg: clay.pillActiveBg, green: clay.pillActiveText,
  greyBg: 'var(--vdms-surface-alt)', grey: 'var(--vdms-text-muted)',
};
const DIFF_STYLE: Record<string, { bg: string; fg: string; label: string }> = {
  new: { bg: C.greenBg, fg: C.green, label: 'New' },
  unchanged: { bg: C.greyBg, fg: C.grey, label: 'Unchanged' },
  reactivate: { bg: C.blueBg, fg: C.blue, label: 'Reactivated' },
  deactivate: { bg: C.amberBg, fg: C.amber, label: 'Deactivated' },
  error: { bg: C.redBg, fg: C.red, label: 'Error' },
};

// Thin wrapper over the shared `dmsBtn` so every call site below (unchanged)
// now renders with Dashboard's button spec instead of this section's own.
const btn = (kind: 'primary' | 'plain' | 'danger', disabled: boolean): React.CSSProperties =>
  dmsBtn(kind === 'plain' ? 'secondary' : kind, !disabled);
const input: React.CSSProperties = dmsControlStyle();
const pill = (bg: string, fg: string): React.CSSProperties => ({ background: bg, color: fg, borderRadius: 999, padding: '2px 8px', fontSize: 11, fontWeight: 600, whiteSpace: 'nowrap' });

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    const d = j && (j.detail !== undefined ? j.detail : j);
    if (typeof d === 'string') return d;
    if (d && typeof d.message === 'string') return d.message;
    return `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

export function TagConfigurationSection({ host }: { host: VesselEmail }): React.ReactElement {
  const [scope, setScope] = React.useState<Scope>('site');
  const [tab, setTab] = React.useState<Tab>('domain');
  const [moreOpen, setMoreOpen] = React.useState(false);
  const [cfg, setCfg] = React.useState<ConfigResponse | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [busy, setBusy] = React.useState('');
  const [error, setError] = React.useState('');
  const [notice, setNotice] = React.useState('');
  const [view, setView] = React.useState<'table' | 'tree'>('table');
  const [search, setSearch] = React.useState('');
  const [fDomain, setFDomain] = React.useState<number | ''>('');
  const [fStatus, setFStatus] = React.useState('');
  const [fSource, setFSource] = React.useState('');
  const [editor, setEditor] = React.useState<{ id: number | null; name: string; code: string; description: string; parentChain: number[] } | null>(null);
  const [panel, setPanel] = React.useState<'' | 'bulk' | 'import' | 'history' | 'sync' | 'copy'>('');
  const [bulkText, setBulkText] = React.useState('');
  const [bulkParent, setBulkParent] = React.useState<number[]>([]);
  const [importRows, setImportRows] = React.useState<ImportRow[] | null>(null);
  const [importMode, setImportMode] = React.useState<Mode>('add');
  const [diff, setDiff] = React.useState<DiffResult | null>(null);
  const [pendingRows, setPendingRows] = React.useState<ImportRow[] | null>(null);
  const [pendingMode, setPendingMode] = React.useState<Mode>('add');
  const [snapshots, setSnapshots] = React.useState<Snapshot[]>([]);
  const [sync, setSync] = React.useState<SyncResult | null>(null);
  const [retagOld, setRetagOld] = React.useState('');
  const [retagNew, setRetagNew] = React.useState('');
  const [retagResult, setRetagResult] = React.useState<RetagResult | null>(null);
  const [vessels, setVessels] = React.useState<{ names: string[]; hint: string } | null>(null);
  const [confirmReset, setConfirmReset] = React.useState(false);
  const [siteOptions, setSiteOptions] = React.useState<SiteOption[]>([]);
  const [activeSiteKey, setActiveSiteKey] = React.useState('');
  const [copySource, setCopySource] = React.useState('');
  const [copyMode, setCopyMode] = React.useState<Mode>('add');
  const [copyDiff, setCopyDiff] = React.useState<CopyDiffResult | null>(null);

  const q = `scope=${scope}`;
  const base = host._base();

  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true); setError('');
    try {
      const r = await fetch(`${base}/api/tag-config?${q}`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      setCfg(await r.json());
    } catch (e) {
      setError((e as Error).message || 'Could not load the tag configuration.');
    } finally {
      setLoading(false);
    }
  }, [host, base, q]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const r = await fetch(`${base}/api/tag-config/sites`, { headers: host._headers() });
        if (!r.ok || cancelled) return;
        const data = await r.json() as { sites: SiteOption[]; active_site: string };
        if (!cancelled) { setSiteOptions(data.sites || []); setActiveSiteKey(data.active_site || ''); }
      } catch {
        // Site picker is a convenience — silently fall back to "current client" only.
      }
    })();
    return () => { cancelled = true; };
  }, [host, base]);

  const call = async (label: string, url: string, init: RequestInit, okMsg?: string): Promise<unknown> => {
    setBusy(label); setError(''); setNotice('');
    try {
      const r = await fetch(url, { ...init, headers: { ...host._headers(), ...(init.headers as Record<string, string> || {}) } });
      if (!r.ok) throw new Error(await readError(r));
      const data = r.status === 204 ? null : await r.json();
      if (okMsg) setNotice(okMsg);
      return data;
    } catch (e) {
      setError((e as Error).message);
      return undefined;
    } finally {
      setBusy('');
    }
  };

  const items = cfg ? cfg.items : [];
  const byId = React.useMemo(() => {
    const m: Record<number, TagItem> = {};
    items.forEach(i => { m[i.id] = i; });
    return m;
  }, [items]);
  const childrenOf = (pid: number | null): TagItem[] =>
    items.filter(i => i.parent_id === pid).sort((a, b) => a.sort_order - b.sort_order || a.name.localeCompare(b.name));
  const domainOf = (it: TagItem): TagItem | undefined => {
    let cur: TagItem | undefined = it;
    let guard = 0;
    while (cur && cur.parent_id !== null && guard++ < 10) cur = byId[cur.parent_id];
    return cur && cur.level === 'domain' ? cur : undefined;
  };
  const isAdmin = !!cfg?.is_admin;
  const level: Level | null = tab === 'vessel_names' || tab === 'internal' ? null : tab;
  // Only Domain, Group and Category have a real SharePoint column to sync
  // (Main Folder has no column; Sub Category is never written to SharePoint).
  const syncableLevel: Level | null = level && SYNCABLE_LEVELS.includes(level) ? level : null;
  const hasSyncWork = !!sync && (!!sync.add.length || !!sync.remove.length || !!sync.term_store?.add.length);
  const mode: Mode = SIMPLE_UI ? 'add' : (level && cfg ? (cfg.modes[level] || 'add') : 'add');
  const internalLevel = SIMPLE_UI && !!level && INTERNAL_LEVELS.includes(level);
  const canEdit = isAdmin && !internalLevel;
  const domains = items.filter(i => i.level === 'domain');

  // ---- cascading parent picker ---------------------------------------
  const ParentPicker = ({ forLevel, chain, onChange }: { forLevel: Level; chain: number[]; onChange: (c: number[]) => void }): React.ReactElement | null => {
    if (forLevel === 'domain') return null;
    const maxIdx = LEVELS.indexOf(forLevel);
    const selects: React.ReactElement[] = [];
    let pid: number | null = null;
    for (let depth = 0; depth < maxIdx; depth++) {
      const opts = (pid === null ? domains : childrenOf(pid)).filter(o => LEVELS.indexOf(o.level) < maxIdx);
      if (!opts.length) break;
      const selected = chain[depth];
      const currentPid = pid;
      selects.push(
        <select key={depth} aria-label={depth === 0 ? 'Domain' : 'Parent level'} style={{ ...input, minWidth: 170 }}
          value={selected === undefined ? '' : String(selected)}
          onChange={e => {
            const v = e.target.value;
            const next = chain.slice(0, depth);
            if (v) next.push(Number(v));
            onChange(next);
          }}>
          <option value="">{depth === 0 ? 'Select Domain…' : `(none — attach to ${currentPid !== null ? byId[currentPid]?.name : ''})`}</option>
          {opts.map(o => (
            <option key={o.id} value={o.id}>{o.name} · {SINGULAR[o.level]}{o.status !== 'Active' ? ` (${o.status})` : ''}</option>
          ))}
        </select>,
      );
      if (selected === undefined) break;
      pid = selected;
    }
    return <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>{selects}</div>;
  };
  const chainFor = (parentId: number | null): number[] => {
    const out: number[] = [];
    let cur = parentId !== null ? byId[parentId] : undefined;
    let guard = 0;
    while (cur && guard++ < 10) { out.unshift(cur.id); cur = cur.parent_id !== null ? byId[cur.parent_id] : undefined; }
    return out;
  };
  const pathOfChain = (chain: number[]): string => chain.map(id => byId[id]?.name || '').join(' > ');

  // ---- filters -------------------------------------------------------
  const rows = items.filter(i => {
    if (!level || i.level !== level) return false;
    if (search && !`${i.name} ${i.code || ''} ${i.path}`.toLowerCase().includes(search.toLowerCase())) return false;
    if (fStatus && i.status !== fStatus) return false;
    if (fSource && i.source !== fSource) return false;
    if (fDomain !== '' && (level === 'domain' ? i.id !== fDomain : domainOf(i)?.id !== fDomain)) return false;
    return true;
  }).sort((a, b) => a.path.localeCompare(b.path));

  // ---- actions -------------------------------------------------------
  const setModeFor = async (lvl: Level, m: Mode): Promise<void> => {
    const res = await call('mode', `${base}/api/tag-config/modes?${q}`, { method: 'PUT', body: JSON.stringify({ modes: { [lvl]: m } }) });
    if (res !== undefined) await load();
  };

  const saveItem = async (): Promise<void> => {
    if (!editor || !level) return;
    const name = editor.name.trim();
    if (!name) { setError('Name is required.'); return; }
    if (INVALID.test(name)) { setError('Name contains characters SharePoint does not allow: ~ " # % * : < > ? / \\ { | }'); return; }
    if (level !== 'domain' && !editor.parentChain.length) { setError(`Select the parent for this ${SINGULAR[level]}.`); return; }
    const parent_id = level === 'domain' ? null : editor.parentChain[editor.parentChain.length - 1];
    const body = { name, code: editor.code || null, description: editor.description || null, parent_id };
    const res = editor.id === null
      ? await call('save', `${base}/api/tag-config/items?${q}`, { method: 'POST', body: JSON.stringify({ ...body, level }) }, 'Saved.')
      : await call('save', `${base}/api/tag-config/items/${editor.id}?${q}`, { method: 'PATCH', body: JSON.stringify(body) }, 'Saved. Existing documents keep their current tag values.');
    if (res !== undefined) { setEditor(null); await load(); }
  };

  const setStatus = async (it: TagItem, status: string): Promise<void> => {
    const res = await call(`status-${it.id}`, `${base}/api/tag-config/items/${it.id}/status?${q}`, { method: 'POST', body: JSON.stringify({ status }) });
    if (res !== undefined) await load();
  };

  const remove = async (it: TagItem): Promise<void> => {
    const res = await call(`del-${it.id}`, `${base}/api/tag-config/items/${it.id}?${q}`, { method: 'DELETE' }, `Deleted '${it.name}'.`);
    if (res !== undefined) await load();
  };

  const move = async (it: TagItem, dir: -1 | 1): Promise<void> => {
    const sibs = childrenOf(it.parent_id).filter(s => s.level === it.level);
    const idx = sibs.findIndex(s => s.id === it.id);
    const j = idx + dir;
    if (idx < 0 || j < 0 || j >= sibs.length) return;
    const ids = sibs.map(s => s.id);
    [ids[idx], ids[j]] = [ids[j], ids[idx]];
    const res = await call(`move-${it.id}`, `${base}/api/tag-config/reorder?${q}`, {
      method: 'POST', body: JSON.stringify({ level: it.level, parent_id: it.parent_id, ordered_ids: ids }),
    });
    if (res !== undefined) await load();
  };

  const runPreview = async (rowsIn: ImportRow[], m: Mode): Promise<void> => {
    const res = await call('preview', `${base}/api/tag-config/bulk/preview?${q}`, { method: 'POST', body: JSON.stringify({ mode: m, rows: rowsIn }) });
    if (res !== undefined) { setDiff(res as DiffResult); setPendingRows(rowsIn); setPendingMode(m); }
  };

  const commit = async (): Promise<void> => {
    if (!pendingRows) return;
    const res = await call('commit', `${base}/api/tag-config/bulk/commit?${q}`, {
      method: 'POST', body: JSON.stringify({ mode: pendingMode, rows: pendingRows, confirm: true, source: panel === 'import' ? 'Imported' : 'Custom' }),
    }) as { created: number; deactivated: number; reactivated: number } | undefined;
    if (res !== undefined) {
      setNotice(`Applied: ${res.created} created, ${res.reactivated} reactivated, ${res.deactivated} deactivated. A snapshot was saved — use Undo to roll back.`);
      setDiff(null); setPendingRows(null); setBulkText(''); setImportRows(null); setPanel('');
      await load();
    }
  };

  const bulkRows = (): ImportRow[] => {
    if (!level) return [];
    const parentPath = level === 'domain' ? '' : pathOfChain(bulkParent);
    return bulkText.split(/\r?\n/).map(s => s.trim()).filter(Boolean).map(name => ({
      Level: SINGULAR[level], 'Parent Path': parentPath, Name: name, Status: 'Active',
    }));
  };

  const parseFile = async (file: File): Promise<void> => {
    setBusy('parse'); setError('');
    try {
      const fd = new FormData();
      fd.append('file', file);
      const h = { ...host._headers() };
      delete h['Content-Type'];
      const r = await fetch(`${base}/api/tag-config/import/parse`, { method: 'POST', headers: h, body: fd });
      if (!r.ok) throw new Error(await readError(r));
      const data = await r.json() as { rows: ImportRow[] };
      setImportRows(data.rows); setDiff(null);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  const download = async (format: 'csv' | 'xlsx'): Promise<void> => {
    setBusy('export'); setError('');
    try {
      const r = await fetch(`${base}/api/tag-config/export?format=${format}&${q}`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      const blob = await r.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `tag-configuration.${format}`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  const loadSnapshots = async (): Promise<void> => {
    const res = await call('snaps', `${base}/api/tag-config/snapshots?${q}`, { method: 'GET' });
    if (res !== undefined) setSnapshots(res as Snapshot[]);
  };
  const restore = async (id: number): Promise<void> => {
    const res = await call('restore', `${base}/api/tag-config/snapshots/${id}/restore?${q}`, { method: 'POST' },
      id === 0 ? 'Last change undone.' : `Restored snapshot #${id}.`);
    if (res !== undefined) { await load(); if (panel === 'history') await loadSnapshots(); }
  };
  const reset = async (): Promise<void> => {
    const res = await call('reset', `${base}/api/tag-config/reset?${q}`, { method: 'POST', body: JSON.stringify({ confirm: true }) },
      'Reset to default values. Custom items were deactivated (not deleted); use Undo to roll back.');
    setConfirmReset(false);
    if (res !== undefined) await load();
  };
  const runCopyPreview = async (): Promise<void> => {
    if (!copySource) { setError('Choose the site to copy from.'); return; }
    const res = await call('copy-preview', `${base}/api/tag-config/copy-from-site/preview?${q}`, {
      method: 'POST', body: JSON.stringify({ source_scope: copySource, mode: copyMode }),
    });
    if (res !== undefined) setCopyDiff(res as CopyDiffResult);
  };

  const copyCommit = async (): Promise<void> => {
    if (!copySource) return;
    const res = await call('copy-commit', `${base}/api/tag-config/copy-from-site/commit?${q}`, {
      method: 'POST', body: JSON.stringify({ source_scope: copySource, mode: copyMode, confirm: true }),
    }) as { created: number; deactivated: number; reactivated: number } | undefined;
    if (res !== undefined) {
      setNotice(`Copied from ${copyDiff?.source_label || copySource}: ${res.created} created, ${res.reactivated} reactivated, ${res.deactivated} deactivated. A snapshot was saved — use Undo to roll back.`);
      setCopyDiff(null); setCopySource(''); setPanel('');
      await load();
    }
  };

  const runSync = async (apply: boolean): Promise<void> => {
    if (!syncableLevel) return;
    const res = await call('sync', `${base}/api/tag-config/domain-sync?${q}`, { method: 'POST', body: JSON.stringify({ apply, level: syncableLevel }) });
    if (res !== undefined) setSync(res as SyncResult);
  };
  // Re-check when the tab changes while the sync panel is open, so it never
  // shows one level's result while labelled for another.
  React.useEffect(() => {
    if (panel !== 'sync') return;
    setSync(null);
    setRetagResult(null);
    if (syncableLevel) runSync(false).catch(() => undefined);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab]);
  const runRetag = async (): Promise<void> => {
    if (!syncableLevel || !retagOld.trim()) return;
    const res = await call('retag', `${base}/api/tag-config/retag?${q}`, {
      method: 'POST',
      body: JSON.stringify({ level: syncableLevel, old_value: retagOld.trim(), new_value: retagNew.trim() || null }),
    });
    if (res !== undefined) setRetagResult(res as RetagResult);
  };
  const loadVessels = async (refresh: boolean): Promise<void> => {
    const res = await call('vessels', `${base}/api/tag-config/vessel-names?${q}&refresh=${refresh}`, { method: 'GET' }) as { vessel_names: string[]; manage_hint: string } | undefined;
    if (res !== undefined) setVessels({ names: res.vessel_names, hint: res.manage_hint });
  };
  React.useEffect(() => { if (tab === 'vessel_names' && !vessels) loadVessels(false).catch(() => undefined); }, [tab]);

  // ---- render helpers ------------------------------------------------
  const tabCount = (t: Tab): string | number =>
    t === 'vessel_names' ? 'Term Store' : t === 'internal' ? 'read-only'
      : items.filter(i => i.level === t && i.status === 'Active').length;
  const internalRows = items.filter(i => INTERNAL_LEVELS.includes(i.level)
    && (!search || `${i.name} ${i.path}`.toLowerCase().includes(search.toLowerCase())))
    .sort((a, b) => a.level.localeCompare(b.level) || a.path.localeCompare(b.path));
  const sharePointLabel = (l: Level): string =>
    l === 'domain' ? 'Column + Term Store' : SYNCABLE_LEVELS.includes(l) ? 'Column choice' : 'Path only';
  const simpleActions = (it: TagItem): React.ReactElement => (
    <div style={{ display: 'flex', gap: 4, justifyContent: 'flex-end' }}>
      <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy}
        onClick={() => setEditor({ id: it.id, name: it.name, code: it.code || '', description: it.description || '', parentChain: chainFor(it.parent_id) })}>Rename</button>
      <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy}
        onClick={() => { setStatus(it, it.status === 'Active' ? 'Inactive' : 'Active').catch(() => undefined); }}>
        {it.status === 'Active' ? 'Archive' : 'Restore'}
      </button>
    </div>
  );
  const statusPill = (s: string): React.ReactElement =>
    s === 'Active' ? <span style={pill(C.greenBg, C.green)}>Active</span> : <span style={pill(C.greyBg, C.grey)}>{s}</span>;

  const actions = (it: TagItem): React.ReactElement => (
    <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
      <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy} aria-label={`Move ${it.name} up`} onClick={() => { move(it, -1).catch(() => undefined); }}>↑</button>
      <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy} aria-label={`Move ${it.name} down`} onClick={() => { move(it, 1).catch(() => undefined); }}>↓</button>
      <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy}
        onClick={() => setEditor({ id: it.id, name: it.name, code: it.code || '', description: it.description || '', parentChain: chainFor(it.parent_id) })}>Edit</button>
      {it.status === 'Active'
        ? <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy} onClick={() => { setStatus(it, 'Inactive').catch(() => undefined); }}>Deactivate</button>
        : <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy} onClick={() => { setStatus(it, 'Active').catch(() => undefined); }}>Activate</button>}
      {!it.is_default && (
        <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy} title="Only unused custom items can be deleted"
          onClick={() => { remove(it).catch(() => undefined); }}>Delete</button>
      )}
    </div>
  );

  const renderTree = (pid: number | null, depth: number): React.ReactElement[] =>
    childrenOf(pid).flatMap(it => {
      const match = !search || `${it.name} ${it.path}`.toLowerCase().includes(search.toLowerCase());
      const own = match ? [(
        <div key={it.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 8px', paddingLeft: 8 + depth * 18, borderBottom: '1px solid var(--vdms-border-soft)', background: it.level === level ? clay.accentSoft : 'transparent' }}>
          <span style={{ fontSize: 13, color: it.status === 'Active' ? C.text : C.sub, fontWeight: it.level === level ? 700 : 400 }}>{it.name}</span>
          <span style={{ fontSize: 11, color: C.sub }}>{SINGULAR[it.level]}</span>
          {statusPill(it.status)}
          {it.level === level && <span style={{ marginLeft: 'auto' }}>{actions(it)}</span>}
        </div>
      )] : [];
      return [...own, ...renderTree(it.id, depth + 1)];
    });

  const diffTable = (d: DiffResult): React.ReactElement => (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 10 }}>
      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
        {(['new', 'unchanged', 'reactivate', 'deactivate', 'error'] as const).map(k => (
          <span key={k} style={pill(DIFF_STYLE[k].bg, DIFF_STYLE[k].fg)}>{DIFF_STYLE[k].label}: {d.counts[k] || 0}</span>
        ))}
      </div>
      {d.mode === 'replace' && (
        <div style={{ padding: '8px 10px', borderRadius: 6, background: C.amberBg, color: C.amber, fontSize: 12 }}>{d.confirmation}</div>
      )}
      <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, maxHeight: 280, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead><tr style={{ background: 'var(--vdms-surface-alt)' }}>
            {['Result', 'Level', 'Name', 'Parent', 'Note'].map(h => <th key={h} style={{ textAlign: 'left', padding: '6px 8px', color: C.sub, borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {d.entries.map((e, i) => (
              <tr key={i} style={{ background: DIFF_STYLE[e.action].bg }}>
                <td style={{ padding: '5px 8px', color: DIFF_STYLE[e.action].fg, fontWeight: 700 }}>{DIFF_STYLE[e.action].label}</td>
                <td style={{ padding: '5px 8px' }}>{e.level}</td>
                <td style={{ padding: '5px 8px' }}>{e.name}</td>
                <td style={{ padding: '5px 8px', color: C.sub }}>{e.parent_path}</td>
                <td style={{ padding: '5px 8px', color: C.sub }}>{e.line ? `Row ${e.line}: ` : ''}{e.message}{e.in_use ? ` · in use by ${e.in_use}` : ''}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );

  const DiffView = (): React.ReactElement | null => {
    if (!diff) return null;
    const hasErrors = (diff.counts.error || 0) > 0;
    return (
      <>
        {diffTable(diff)}
        <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
          <button style={btn('primary', hasErrors || !!busy)} disabled={hasErrors || !!busy} onClick={() => { commit().catch(() => undefined); }}>
            {busy === 'commit' ? 'Applying…' : 'Confirm and apply'}
          </button>
          <button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => { setDiff(null); setPendingRows(null); }}>Cancel</button>
          {hasErrors && <span style={{ fontSize: 12, color: C.red, alignSelf: 'center' }}>Fix the errors first — nothing has been saved.</span>}
        </div>
      </>
    );
  };

  if (loading && !cfg) return <div style={{ fontSize: 13, color: C.sub }}>Loading tag configuration…</div>;
  if (!cfg) {
    return (
      <div style={{ padding: 12, borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>
        {error || 'Tag configuration is unavailable.'}{' '}
        <button style={btn('plain', false)} onClick={() => { load().catch(() => undefined); }}>Retry</button>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, maxWidth: 1040 }}>
      {SIMPLE_UI && (
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 8, flexWrap: 'wrap' }}>
          <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>Tag configuration</div>
          <div style={{ fontSize: 12, color: C.sub }}>Changes are pushed to SharePoint and the Term Store automatically.</div>
        </div>
      )}
      {/* Header: scope + global actions */}
      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
        {!SIMPLE_UI && <div style={{ fontSize: 12, color: C.sub }}>
          Client: <b style={{ color: C.text }}>
            {scope === 'template'
              ? 'Default template for new clients'
              : (siteOptions.find(s => s.site_key === cfg.site_key)?.display_name || cfg.site_key)}
          </b>
          {cfg.origin !== 'site' && scope !== 'template' && <span> · using defaults until the first change</span>}
        </div>}
        {!SIMPLE_UI && <select aria-label="Scope" style={input} value={scope} onChange={e => { setScope(e.target.value); setDiff(null); setEditor(null); setSync(null); }}>
          <option value="site">Apply to current client{activeSiteKey ? ` (${siteOptions.find(s => s.site_key === activeSiteKey)?.display_name || activeSiteKey})` : ''}</option>
          {siteOptions.filter(s => s.site_key !== activeSiteKey).map(s => (
            <option key={s.site_key} value={s.site_key}>{s.display_name}</option>
          ))}
          <option value="template" disabled={!isAdmin}>Default template for new clients (Admin)</option>
        </select>}
        <div style={{ marginLeft: 'auto', position: 'relative' }}>
          <button type="button" aria-haspopup="menu" aria-expanded={moreOpen} style={btn('plain', false)} onClick={() => setMoreOpen(o => !o)}>More ▾</button>
          {moreOpen && (
          <div role="menu" className="vdms-more-menu" onClick={() => setMoreOpen(false)} style={{ position: 'absolute', right: 0, top: 'calc(100% + 6px)', zIndex: 20, minWidth: 230, display: 'flex', flexDirection: 'column', gap: 2, padding: 6, borderRadius: 12, background: 'var(--vdms-surface)', border: '1px solid var(--vdms-line)', boxShadow: '0 12px 28px rgba(16,40,80,0.18)' }}>
          <style>{`.vdms-more-menu button { width: 100%; text-align: left !important; border: none !important; background: transparent !important; box-shadow: none !important; } .vdms-more-menu button:hover:not(:disabled) { background: var(--vdms-surface-alt) !important; }`}</style>
          <button style={btn('plain', !isAdmin || !!busy)} disabled={!isAdmin || !!busy} onClick={() => { restore(0).catch(() => undefined); }}>Undo last change</button>
          <button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => { setPanel(panel === 'history' ? '' : 'history'); loadSnapshots().catch(() => undefined); }}>History</button>
          <button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => { download('csv').catch(() => undefined); }}>Export CSV</button>
          <button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => { download('xlsx').catch(() => undefined); }}>Export Excel</button>
          <button style={btn('plain', !isAdmin || !!busy)} disabled={!isAdmin || !!busy} onClick={() => { setPanel(panel === 'import' ? '' : 'import'); setImportRows(null); setDiff(null); }}>Bulk import</button>
          {!SIMPLE_UI && <button style={btn('plain', !isAdmin || !!busy)} disabled={!isAdmin || !!busy} onClick={() => { setPanel(panel === 'copy' ? '' : 'copy'); setCopyDiff(null); }}>Copy from other site</button>}
          <button style={btn('plain', !!busy || scope === 'template' || !syncableLevel)} disabled={!!busy || scope === 'template' || !syncableLevel}
            title={syncableLevel ? '' : 'Switch to the Domains, Groups or Categories tab to sync its SharePoint column'}
            onClick={() => { setPanel(panel === 'sync' ? '' : 'sync'); setSync(null); runSync(false).catch(() => undefined); }}>SharePoint tag sync</button>
          {!SIMPLE_UI && <button style={btn('danger', !isAdmin || !!busy)} disabled={!isAdmin || !!busy} onClick={() => setConfirmReset(true)}>Reset to Default</button>}
          </div>
          )}
        </div>
      </div>

      {!isAdmin && <div style={{ fontSize: 12, color: C.sub }}>Read-only: only administrators can change the tag configuration.</div>}
      {error && <div role="alert" style={{ padding: '8px 12px', borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>{error}</div>}
      {notice && <div role="status" style={{ padding: '8px 12px', borderRadius: 8, background: C.greenBg, color: C.green, fontSize: 13 }}>{notice}</div>}

      {confirmReset && (
        <div style={{ padding: 12, borderRadius: 8, border: `1px solid ${C.red}`, background: C.redBg, fontSize: 13, color: C.text }}>
          Reset every level to the default values (including the four default Domains)? Custom items are deactivated, not deleted. A snapshot is saved first so you can undo.
          <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
            <button style={btn('danger', !!busy)} disabled={!!busy} onClick={() => { reset().catch(() => undefined); }}>{busy === 'reset' ? 'Resetting…' : 'Yes, reset'}</button>
            <button style={btn('plain', false)} onClick={() => setConfirmReset(false)}>Cancel</button>
          </div>
        </div>
      )}

      {panel === 'history' && (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 6 }}>Snapshots (saved before every Replace, Import, Reset and Restore)</div>
          {!snapshots.length ? <div style={{ fontSize: 12, color: C.sub }}>No snapshots yet.</div> : (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <tbody>
                {snapshots.map(s => (
                  <tr key={s.id} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                    <td style={{ padding: '4px 6px' }}>#{s.id}</td>
                    <td style={{ padding: '4px 6px' }}>{s.created_at ? new Date(s.created_at).toLocaleString() : ''}</td>
                    <td style={{ padding: '4px 6px' }}>{s.mode}</td>
                    <td style={{ padding: '4px 6px', color: C.sub }}>{s.changed_by}</td>
                    <td style={{ padding: '4px 6px', textAlign: 'right' }}>
                      <button style={btn('plain', !isAdmin || !!busy)} disabled={!isAdmin || !!busy} onClick={() => { restore(s.id).catch(() => undefined); }}>Restore this version</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}

      {panel === 'sync' && (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, fontSize: 13 }}>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>SharePoint {syncableLevel ? SINGULAR[syncableLevel] : ''} column</div>
          {!syncableLevel ? (
            <div style={{ color: C.sub }}>Switch to the Domains, Groups or Categories tab to sync its SharePoint column. Main Folder has no column of its own, and Sub Category is never written to SharePoint.</div>
          ) : !sync ? <div style={{ color: C.sub }}>Checking…</div> : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <div>Column: <b>{sync.column?.displayName || '—'}</b> · type: <b>{sync.type || '—'}</b></div>
              {sync.error && <div style={{ color: C.red }}>{sync.error}</div>}
              {sync.type === 'text' && <div style={{ color: C.sub }}>Plain text column — nothing to sync.</div>}
              {!!sync.add.length && <div>To add: {sync.add.join(', ')}</div>}
              {!!sync.remove.length && <div>To remove from choices (existing documents keep their value): {sync.remove.join(', ')}</div>}
              {!sync.error && !sync.add.length && !sync.remove.length && sync.type !== 'text' && <div style={{ color: C.green }}>In sync with the Active {LABEL[syncableLevel]}.</div>}
              {sync.applied && <div style={{ color: C.green }}>Updated.</div>}
              {sync.term_store && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {sync.term_store.error && <div style={{ color: C.red }}>Term Store: {sync.term_store.error}</div>}
                  {!!sync.term_store.added.length && <div style={{ color: C.green }}>Term Store: created {sync.term_store.added.join(', ')}.</div>}
                  {!sync.term_store.added.length && !!sync.term_store.add.length && <div>Term Store — term sets to create in "Vessel DMS": {sync.term_store.add.join(', ')}</div>}
                  {!sync.term_store.error && !sync.term_store.add.length && !sync.term_store.added.length && <div style={{ color: C.green }}>Term Store has a term set for every Active Domain.</div>}
                </div>
              )}
              <div>
                <button style={btn('primary', !isAdmin || !!busy || !hasSyncWork)}
                  disabled={!isAdmin || !!busy || !hasSyncWork}
                  onClick={() => { runSync(true).catch(() => undefined); }}>Sync now</button>
              </div>

              <div style={{ borderTop: `1px solid ${C.border}`, marginTop: 4, paddingTop: 10 }}>
                <div style={{ fontWeight: 700, marginBottom: 4 }}>Retag existing documents</div>
                <div style={{ color: C.sub, fontSize: 12, marginBottom: 8 }}>
                  A rename or a Deactivate now does this automatically for the item involved. Use this to
                  fix documents left over from before that (e.g. a value renamed or deactivated earlier),
                  or to re-run one that reported errors below. Leave "New value" blank to clear the field
                  instead of replacing it.
                </div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                  <input aria-label="Old value" style={{ ...input, minWidth: 180 }} placeholder="Old value on the documents"
                    value={retagOld} onChange={e => setRetagOld(e.target.value)} />
                  <input aria-label="New value" style={{ ...input, minWidth: 180 }} placeholder="New value (blank = clear)"
                    value={retagNew} onChange={e => setRetagNew(e.target.value)} />
                  <button style={btn('primary', !isAdmin || !!busy || !retagOld.trim())}
                    disabled={!isAdmin || !!busy || !retagOld.trim()}
                    onClick={() => { runRetag().catch(() => undefined); }}>
                    {busy === 'retag' ? 'Retagging…' : 'Retag now'}
                  </button>
                </div>
                {retagResult && (
                  <div style={{ marginTop: 8 }}>
                    <div>Checked {retagResult.checked} document(s), updated {retagResult.updated}.</div>
                    {!!retagResult.errors.length && (
                      <div style={{ color: C.red }}>
                        {retagResult.errors.length} error(s): {retagResult.errors.slice(0, 5).join('; ')}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {panel === 'copy' && (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, fontSize: 13 }}>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>Copy from another site</div>
          <div style={{ color: C.sub, fontSize: 12, marginBottom: 8 }}>
            Pulls another client's Active Domains, Main Folders, Groups, Categories and Sub Categories onto{' '}
            <b>{cfg.site_key === 'template' || cfg.scope === 'template' ? 'the default template' : (siteOptions.find(s => s.site_key === cfg.site_key)?.display_name || cfg.site_key)}</b>.
            Use this for a site that has no Domain/tags of its own yet, instead of leaving it on the generic template.
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center', marginBottom: 8 }}>
            <select aria-label="Copy from" style={input} value={copySource} onChange={e => { setCopySource(e.target.value); setCopyDiff(null); }}>
              <option value="">Select a site…</option>
              {siteOptions.filter(s => s.site_key !== cfg.site_key).map(s => (
                <option key={s.site_key} value={s.site_key}>{s.display_name}{s.site_key === activeSiteKey ? ' (current client)' : ''}</option>
              ))}
              {cfg.scope !== 'template' && <option value="template">Default template for new clients</option>}
            </select>
            <label style={{ marginRight: 4 }}><input type="radio" name="copy-mode" checked={copyMode === 'add'} onChange={() => { setCopyMode('add'); setCopyDiff(null); }} /> Add to existing</label>
            <label><input type="radio" name="copy-mode" checked={copyMode === 'replace'} onChange={() => { setCopyMode('replace'); setCopyDiff(null); }} /> Replace existing</label>
            <button style={btn('plain', !copySource || !!busy)} disabled={!copySource || !!busy} onClick={() => { runCopyPreview().catch(() => undefined); }}>
              {busy === 'copy-preview' ? 'Comparing…' : 'Preview'}
            </button>
          </div>
          {copyDiff && (
            <>
              {diffTable(copyDiff)}
              <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                <button style={btn('primary', (copyDiff.counts.error || 0) > 0 || !!busy)} disabled={(copyDiff.counts.error || 0) > 0 || !!busy}
                  onClick={() => { copyCommit().catch(() => undefined); }}>
                  {busy === 'copy-commit' ? 'Applying…' : 'Confirm and apply'}
                </button>
                <button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => setCopyDiff(null)}>Cancel</button>
                {(copyDiff.counts.error || 0) > 0 && <span style={{ fontSize: 12, color: C.red, alignSelf: 'center' }}>Fix the errors first — nothing has been saved.</span>}
              </div>
            </>
          )}
        </div>
      )}

      {panel === 'import' && (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, fontSize: 13 }}>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>Bulk import</div>
          <div style={{ color: C.sub, fontSize: 12, marginBottom: 8 }}>
            Columns: Level, Parent Path (e.g. "Technical &amp; Crewing &gt; Drawings and Manuals &gt; Drawings"), Name, Code, Description, Sort Order, Status. An exported file can be re-imported.
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            <div><b>Step 1</b> Upload file:{' '}
              <input type="file" accept=".csv,.xlsx" aria-label="Import file"
                onChange={e => { const f = e.target.files && e.target.files[0]; if (f) parseFile(f).catch(() => undefined); }} />
              {importRows && <span style={{ marginLeft: 8, color: C.sub }}>{importRows.length} rows read</span>}
            </div>
            {importRows && (
              <div><b>Step 2</b> Mode:{' '}
                <label style={{ marginRight: 12 }}><input type="radio" name="import-mode" checked={importMode === 'add'} onChange={() => setImportMode('add')} /> Add to existing</label>
                <label><input type="radio" name="import-mode" checked={importMode === 'replace'} onChange={() => setImportMode('replace')} /> Replace existing</label>
              </div>
            )}
            {importRows && (
              <div><b>Step 3</b>{' '}
                <button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => { runPreview(importRows, importMode).catch(() => undefined); }}>
                  {busy === 'preview' ? 'Checking…' : 'Preview changes'}
                </button>
              </div>
            )}
            {diff && pendingRows === importRows && <div><b>Step 4</b> Review and confirm. Nothing is saved until you confirm.<DiffView /></div>}
          </div>
        </div>
      )}

      {/* Level tabs */}
      <div role="tablist" aria-label="Tag levels" style={{ display: 'flex', gap: 4, borderBottom: `1px solid ${C.border}`, flexWrap: 'wrap' }}>
        {(SIMPLE_UI ? SIMPLE_TABS : [...LEVELS, 'vessel_names'] as Tab[]).map(t => (
          <button key={t} role="tab" aria-selected={tab === t}
            onClick={() => { setTab(t); setEditor(null); setDiff(null); setPanel(panel === 'import' || panel === 'history' || panel === 'sync' ? panel : ''); setBulkText(''); setBulkParent([]); }}
            style={{ border: 'none', borderBottom: `2px solid ${tab === t ? C.blue : 'transparent'}`, background: 'transparent', padding: '8px 12px', fontSize: 13, fontWeight: tab === t ? 700 : 500, color: tab === t ? C.blue : C.grey, cursor: 'pointer' }}>
            {t === 'internal' ? '\u{1F512} Internal' : LABEL[t]}
            {(SIMPLE_UI || t !== 'vessel_names') && <span style={{ marginLeft: 6, fontSize: 11, color: C.sub }}>{SIMPLE_UI ? tabCount(t) : items.filter(i => i.level === t && i.status === 'Active').length}</span>}
          </button>
        ))}
      </div>

      {tab === 'internal' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, fontSize: 13 }}>
          <div style={{ padding: '8px 12px', borderRadius: 8, background: C.blueBg, color: C.text }}>
            Main Folders and Sub Categories only help read folder paths. They are not SharePoint columns, so they are view-only here.
          </div>
          <input aria-label="Search" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} style={{ ...input, width: 240 }} />
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'auto', maxHeight: 520 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: 'var(--vdms-surface-alt)', position: 'sticky', top: 0 }}>
                  {['Name', 'Type', 'Used for'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '7px 8px', color: C.sub, fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {!internalRows.length && <tr><td colSpan={3} style={{ padding: 14, color: C.sub }}>Nothing matches your search.</td></tr>}
                {internalRows.map(it => (
                  <tr key={it.id} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                    <td style={{ padding: '6px 8px', fontWeight: 600 }}>{it.name}</td>
                    <td style={{ padding: '6px 8px', color: C.sub }}>{SINGULAR[it.level]}</td>
                    <td style={{ padding: '6px 8px', color: C.sub }}>Path parsing</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === 'vessel_names' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, fontSize: 13 }}>
          <div style={{ padding: '8px 12px', borderRadius: 8, background: C.blueBg, color: C.text }}>
            Vessel Names come only from the SharePoint Term Store and cannot be edited here. {vessels?.hint}
          </div>
          <div><button style={btn('plain', !!busy)} disabled={!!busy} onClick={() => { loadVessels(true).catch(() => undefined); }}>{busy === 'vessels' ? 'Refreshing…' : 'Refresh / Sync from Term Store'}</button></div>
          {vessels && (
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, maxHeight: 360, overflow: 'auto' }}>
              {!vessels.names.length ? <div style={{ padding: 10, color: C.sub }}>No vessel terms found.</div>
                : vessels.names.map(n => <div key={n} style={{ padding: '5px 10px', borderBottom: '1px solid var(--vdms-border-soft)' }}>{n}</div>)}
            </div>
          )}
        </div>
      )}

      {level && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {/* Customisation mode for this level */}
          {internalLevel && (
            <div style={{ padding: '8px 12px', borderRadius: 8, background: C.blueBg, color: C.text, fontSize: 13 }}>
              {LABEL[level]} only help read folder paths. They are not SharePoint columns, so they are view-only here.
            </div>
          )}
          {!SIMPLE_UI && <div role="radiogroup" aria-label={`Customisation mode for ${LABEL[level]}`} style={{ display: 'flex', gap: 14, alignItems: 'center', flexWrap: 'wrap', fontSize: 13 }}>
            <b>Customisation mode:</b>
            <label><input type="radio" name={`mode-${level}`} checked={mode === 'add'} disabled={!isAdmin || !!busy} onChange={() => { setModeFor(level, 'add').catch(() => undefined); }} /> Add to existing</label>
            <label><input type="radio" name={`mode-${level}`} checked={mode === 'replace'} disabled={!isAdmin || !!busy} onChange={() => { setModeFor(level, 'replace').catch(() => undefined); }} /> Replace existing</label>
            <span style={{ fontSize: 12, color: C.sub }}>
              {mode === 'add' ? 'Defaults stay; new values are added alongside them.'
                : 'Your list replaces the current values under the chosen parent. Replaced items are deactivated (never deleted) so existing documents keep their tags.'}
            </span>
          </div>}

          {/* Toolbar */}
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            <input aria-label="Search" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} style={{ ...input, width: 200 }} />
            {!SIMPLE_UI && level !== 'domain' && (
              <select aria-label="Filter by Domain" style={input} value={fDomain === '' ? '' : String(fDomain)} onChange={e => setFDomain(e.target.value ? Number(e.target.value) : '')}>
                <option value="">All Domains</option>
                {domains.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
              </select>
            )}
            {!SIMPLE_UI && <select aria-label="Filter by status" style={input} value={fStatus} onChange={e => setFStatus(e.target.value)}>
              <option value="">All statuses</option><option>Active</option><option>Inactive</option><option>Archived</option>
            </select>}
            {!SIMPLE_UI && <select aria-label="Filter by source" style={input} value={fSource} onChange={e => setFSource(e.target.value)}>
              <option value="">All sources</option><option>Default</option><option>Custom</option><option>Imported</option>
            </select>}
            {!SIMPLE_UI && <div role="group" aria-label="View" style={{ display: 'flex' }}>
              {(['table', 'tree'] as const).map(v => (
                <button key={v} onClick={() => setView(v)} aria-pressed={view === v}
                  style={{ ...btn('plain', false), background: view === v ? C.blueBg : 'var(--vdms-surface)', color: view === v ? C.blue : C.text, borderRadius: v === 'table' ? '6px 0 0 6px' : '0 6px 6px 0' }}>
                  {v === 'table' ? 'List' : 'Tree'}
                </button>
              ))}
            </div>}
            {!internalLevel && <div style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
              <button style={btn('primary', !canEdit || !!busy)} disabled={!canEdit || !!busy}
                onClick={() => { setEditor({ id: null, name: '', code: '', description: '', parentChain: [] }); setPanel(''); }}>+ Add {SINGULAR[level]}</button>
              <button style={btn('plain', !canEdit || !!busy)} disabled={!canEdit || !!busy}
                onClick={() => { setPanel(panel === 'bulk' ? '' : 'bulk'); setEditor(null); setDiff(null); }}>
                {mode === 'replace' ? 'Replace list…' : 'Add several…'}
              </button>
            </div>}
          </div>

          {/* Add / edit form */}
          {editor && (
            <div style={{ border: `1px solid ${C.blue}`, borderRadius: 8, padding: 12, background: clay.accentSoft, display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ fontSize: 13, fontWeight: 700 }}>{editor.id === null ? `New ${SINGULAR[level]}` : `Edit ${SINGULAR[level]}`}</div>
              {level !== 'domain' && (
                <div style={{ fontSize: 12 }}>Parent{editor.id === null ? ' (select from Domain down; levels may be skipped)' : ''}:
                  <ParentPicker forLevel={level} chain={editor.parentChain} onChange={c => setEditor({ ...editor, parentChain: c })} />
                </div>
              )}
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <input aria-label="Name" placeholder="Name *" maxLength={cfg.folder_name_rules.max_length} value={editor.name}
                  onChange={e => setEditor({ ...editor, name: e.target.value })} style={{ ...input, minWidth: 260 }} />
                {!SIMPLE_UI && <input aria-label="Code" placeholder="Code (optional)" maxLength={40} value={editor.code}
                  onChange={e => setEditor({ ...editor, code: e.target.value })} style={{ ...input, width: 140 }} />}
                {!SIMPLE_UI && <input aria-label="Description" placeholder="Description (optional)" value={editor.description}
                  onChange={e => setEditor({ ...editor, description: e.target.value })} style={{ ...input, flex: 1, minWidth: 200 }} />}
              </div>
              <div style={{ fontSize: 11, color: C.sub }}>
                "&amp;" is allowed in names; the SharePoint folder name maps &amp; → and. Not allowed: {cfg.folder_name_rules.invalid_chars}. Max {cfg.folder_name_rules.max_length} characters.
                {editor.id !== null && ' Renaming keeps the old name as an alias — documents already tagged keep their value.'}
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button style={btn('primary', !!busy)} disabled={!!busy} onClick={() => { saveItem().catch(() => undefined); }}>{busy === 'save' ? 'Saving…' : 'Save'}</button>
                <button style={btn('plain', false)} onClick={() => setEditor(null)}>Cancel</button>
              </div>
            </div>
          )}

          {/* Bulk add / replace list */}
          {panel === 'bulk' && (
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ fontSize: 13, fontWeight: 700 }}>
                {mode === 'replace' ? `Replace ${LABEL[level]}` : `Add ${LABEL[level]}`} — one name per line
              </div>
              {level !== 'domain' && (
                <div style={{ fontSize: 12 }}>Parent: <ParentPicker forLevel={level} chain={bulkParent} onChange={setBulkParent} /></div>
              )}
              <textarea aria-label="Names, one per line" rows={6} value={bulkText} onChange={e => setBulkText(e.target.value)} style={{ ...input, fontFamily: 'inherit' }} />
              <div>
                <button style={btn('plain', !!busy || !bulkText.trim() || (level !== 'domain' && !bulkParent.length))}
                  disabled={!!busy || !bulkText.trim() || (level !== 'domain' && !bulkParent.length)}
                  onClick={() => { runPreview(bulkRows(), mode).catch(() => undefined); }}>
                  {busy === 'preview' ? 'Checking…' : 'Preview changes'}
                </button>
              </div>
              {diff && pendingRows !== importRows && <DiffView />}
            </div>
          )}

          {/* List / tree */}
          {SIMPLE_UI || view === 'table' ? (
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'auto', maxHeight: 520 }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: 'var(--vdms-surface-alt)', position: 'sticky', top: 0 }}>
                    {(SIMPLE_UI ? ['Name', 'Status', 'In SharePoint', 'Actions'] : ['Name', 'Parent path', 'Code', 'Order', 'Status', 'Source', 'Folder name', '']).map(h => (
                      <th key={h} style={{ textAlign: 'left', padding: '7px 8px', color: C.sub, fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {!rows.length && (
                    <tr><td colSpan={SIMPLE_UI ? 4 : 8} style={{ padding: 14, color: C.sub }}>No {LABEL[level].toLowerCase()} match these filters.</td></tr>
                  )}
                  {rows.map(it => (
                    <tr key={it.id} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                      <td style={{ padding: '6px 8px', color: it.status === 'Active' ? C.text : C.sub, fontWeight: 600 }} title={it.description || ''}>
                        {it.name}
                        {SIMPLE_UI && it.path.includes(' > ') && <div style={{ fontSize: 11, fontWeight: 400, color: C.sub }}>{it.path.split(' > ').slice(0, -1).join(' > ')}</div>}
                      </td>
                      {!SIMPLE_UI && <td style={{ padding: '6px 8px', color: C.sub }}>{it.path.split(' > ').slice(0, -1).join(' > ') || '—'}</td>}
                      {!SIMPLE_UI && <td style={{ padding: '6px 8px' }}>{it.code || ''}</td>}
                      {!SIMPLE_UI && <td style={{ padding: '6px 8px' }}>{it.sort_order}</td>}
                      <td style={{ padding: '6px 8px' }}>{statusPill(it.status)}</td>
                      {SIMPLE_UI && <td style={{ padding: '6px 8px', color: C.sub }}>{sharePointLabel(it.level)}</td>}
                      {!SIMPLE_UI && <td style={{ padding: '6px 8px', color: C.sub }}>{it.source}</td>}
                      {!SIMPLE_UI && <td style={{ padding: '6px 8px', color: C.sub }}>{it.folder_name}</td>}
                      <td style={{ padding: '6px 8px', textAlign: SIMPLE_UI ? 'right' : 'left' }}>{SIMPLE_UI ? simpleActions(it) : actions(it)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'auto', maxHeight: 520 }}>
              {renderTree(null, 0)}
            </div>
          )}
        </div>
      )}

      {SIMPLE_UI && (
        <div style={{ marginTop: 4, paddingTop: 12, borderTop: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
          <div style={{ fontSize: 12, color: C.sub }}>Renaming keeps the old name, so existing files still match.</div>
          <button style={btn('plain', !isAdmin || !!busy)} disabled={!isAdmin || !!busy}
            title="Fix documents tagged with an old value. Opens at Domains, Groups or Categories."
            onClick={() => {
              setPanel('sync'); setSync(null); setRetagResult(null);
              if (syncableLevel) runSync(false).catch(() => undefined); else setTab('domain');
              window.scrollTo({ top: 0, behavior: 'smooth' });
            }}>
            Retag existing files (admin)
          </button>
        </div>
      )}
    </div>
  );
}
