import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';

/**
 * Pick files and folders from any SharePoint site in Site Management, for
 * the email composer. Browses with the same backend folder listing the
 * Documents page uses (host._loadSiteFolderChildren).
 */

export interface PickedItem {
  drive_id: string;
  item_id: string;
  name: string;
  is_folder: boolean;
  size: number;
  site_name: string;
  mode: 'attach' | 'link';
}

interface Crumb { id: string; name: string }

const LINE = 'var(--vdms-line)';

export function formatSize(n: number): string {
  if (!n) return '';
  const u = ['B', 'KB', 'MB', 'GB'];
  const i = Math.min(u.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
  return `${(n / Math.pow(1024, i)).toFixed(i >= 2 ? 1 : 0)} ${u[i]}`;
}

export function fileIcon(name: string, isFolder: boolean): { icon: string; color: string } {
  if (isFolder) return { icon: 'FabricFolderFill', color: '#d4a017' };
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (ext === 'pdf') return { icon: 'PDF', color: '#dc2626' };
  if (['doc', 'docx'].indexOf(ext) >= 0) return { icon: 'WordDocument', color: '#2563eb' };
  if (['xls', 'xlsx', 'csv'].indexOf(ext) >= 0) return { icon: 'ExcelDocument', color: '#16a34a' };
  if (['ppt', 'pptx'].indexOf(ext) >= 0) return { icon: 'PowerPointDocument', color: '#ea580c' };
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff'].indexOf(ext) >= 0) return { icon: 'FileImage', color: '#9333ea' };
  if (['dwg', 'dxf'].indexOf(ext) >= 0) return { icon: 'FileCAD', color: '#0e7490' };
  return { icon: 'Page', color: '#64748b' };
}

export function SharePointPicker(props: {
  host: VesselEmail;
  onAdd: (items: PickedItem[]) => void;
  onClose: () => void;
}): React.ReactElement {
  const { host } = props;
  const sites = React.useMemo(() => {
    const seen = new Set<string>();
    return (host.state.documentSites || []).filter(s => {
      const k = String(s.drive_id || '');
      if (!k || !s.site_id || seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }, [host.state.documentSites]);

  const [siteIdx, setSiteIdx] = React.useState(0);
  const [crumbs, setCrumbs] = React.useState<Crumb[]>([{ id: 'root', name: 'Documents' }]);
  const [items, setItems] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [filter, setFilter] = React.useState('');
  const [selected, setSelected] = React.useState<Record<string, PickedItem>>({});

  const site = sites[siteIdx];
  const folderId = crumbs[crumbs.length - 1].id;

  React.useEffect(() => {
    if (!site) return;
    let cancelled = false;
    setLoading(true); setError(''); setFilter('');
    host._loadSiteFolderChildren(site.site_id, site.drive_id, folderId)
      .then(res => {
        if (cancelled) return;
        if (res.error) setError(res.throttled ? 'SharePoint is busy — try again in a moment.' : "Couldn't load this folder.");
        const list = (res.items || []).slice().sort((a: any, b: any) =>
          (b.folder ? 1 : 0) - (a.folder ? 1 : 0) || String(a.name).localeCompare(String(b.name)));
        setItems(list);
      })
      .catch(() => { if (!cancelled) setError("Couldn't load this folder."); })
      .then(() => { if (!cancelled) setLoading(false); }, () => undefined);
    return () => { cancelled = true; };
  }, [site && site.drive_id, folderId]);

  const toggle = (it: any): void => {
    if (!site) return;
    setSelected(prev => {
      const next = { ...prev };
      if (next[it.id]) delete next[it.id];
      else next[it.id] = {
        drive_id: site.drive_id, item_id: it.id, name: it.name, is_folder: !!it.folder,
        size: Number(it.size || 0), site_name: site.sp_site_name || site.site_key,
        mode: it.folder ? 'link' : 'attach',
      };
      return next;
    });
  };

  const q = filter.trim().toLowerCase();
  const shown = q ? items.filter(i => String(i.name || '').toLowerCase().indexOf(q) >= 0) : items;
  const picked = Object.keys(selected).map(k => selected[k]);

  return (
    <div style={{ position: 'absolute', inset: 0, background: 'rgba(11,42,74,0.35)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2, padding: 16 }}
      onClick={e => { if (e.target === e.currentTarget) props.onClose(); }}>
      <div style={{ width: 640, maxWidth: '100%', height: '82%', background: clay.surface, borderRadius: 14, border: `1px solid ${LINE}`, boxShadow: '0 18px 48px rgba(11,42,74,0.3)', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
        <div style={{ padding: '12px 16px', borderBottom: `1px solid ${LINE}`, display: 'flex', alignItems: 'center', gap: 10 }}>
          <Icon iconName="SharepointLogo" style={{ fontSize: 18, color: '#0e7490' }} />
          <div style={{ fontWeight: 800, fontSize: 15, color: clay.text, flex: 1 }}>Attach from SharePoint</div>
          <button type="button" onClick={props.onClose} aria-label="Close" style={{ border: 'none', background: 'none', cursor: 'pointer', color: clay.textMuted }}>
            <Icon iconName="Cancel" />
          </button>
        </div>

        <div style={{ padding: '10px 16px', display: 'flex', gap: 8, borderBottom: `1px solid ${LINE}`, flexWrap: 'wrap' }}>
          <select value={siteIdx} onChange={e => { setSiteIdx(Number(e.target.value)); setCrumbs([{ id: 'root', name: 'Documents' }]); }}
            style={{ flex: '1 1 200px', padding: '7px 8px', borderRadius: 8, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text, fontSize: 13 }}>
            {sites.map((s, i) => <option key={s.drive_id} value={i}>{s.sp_site_name || s.site_key}</option>)}
          </select>
          <input value={filter} onChange={e => setFilter(e.target.value)} placeholder="Filter this folder…"
            style={{ flex: '1 1 180px', padding: '7px 10px', borderRadius: 8, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text, fontSize: 13 }} />
        </div>

        <div style={{ padding: '8px 16px', fontSize: 12.5, color: clay.textMuted, display: 'flex', flexWrap: 'wrap', gap: 4, alignItems: 'center' }}>
          {crumbs.map((c, i) => (
            <React.Fragment key={c.id + i}>
              {i > 0 && <Icon iconName="ChevronRight" style={{ fontSize: 9 }} />}
              <button type="button" onClick={() => setCrumbs(crumbs.slice(0, i + 1))} disabled={i === crumbs.length - 1}
                style={{ border: 'none', background: 'none', padding: 0, cursor: i === crumbs.length - 1 ? 'default' : 'pointer', color: i === crumbs.length - 1 ? clay.text : clay.accent, fontWeight: i === crumbs.length - 1 ? 700 : 600, fontSize: 12.5 }}>
                {c.name}
              </button>
            </React.Fragment>
          ))}
        </div>

        <div style={{ flex: 1, overflowY: 'auto', padding: '0 8px 8px' }}>
          {!sites.length && <div style={{ padding: 20, color: clay.textMuted, fontSize: 13 }}>No SharePoint sites are configured yet (Sites → Site Management).</div>}
          {loading && <div style={{ padding: 20, color: clay.textMuted, fontSize: 13 }}>Loading…</div>}
          {!loading && error && <div style={{ padding: 20, color: '#b91c1c', fontSize: 13 }}>{error}</div>}
          {!loading && !error && shown.length === 0 && <div style={{ padding: 20, color: clay.textMuted, fontSize: 13 }}>This folder is empty.</div>}
          {!loading && shown.map(it => {
            const fi = fileIcon(it.name, !!it.folder);
            const on = !!selected[it.id];
            return (
              <div key={it.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '7px 8px', borderRadius: 8, background: on ? clay.accentSoft : 'transparent' }}>
                <input type="checkbox" checked={on} onChange={() => toggle(it)} aria-label={`Select ${it.name}`} style={{ width: 16, height: 16, cursor: 'pointer' }} />
                <Icon iconName={fi.icon} style={{ fontSize: 17, color: fi.color }} />
                <button type="button"
                  onClick={() => (it.folder ? setCrumbs([...crumbs, { id: it.id, name: it.name }]) : toggle(it))}
                  style={{ flex: 1, minWidth: 0, textAlign: 'left', border: 'none', background: 'none', padding: 0, cursor: 'pointer', color: clay.text, fontSize: 13, fontWeight: it.folder ? 600 : 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {it.name}
                </button>
                <span style={{ fontSize: 11.5, color: clay.textMuted, whiteSpace: 'nowrap' }}>
                  {it.folder ? `${(it.folder.childCount ?? '')} ${it.folder.childCount === 1 ? 'item' : it.folder.childCount !== undefined ? 'items' : ''}` : formatSize(Number(it.size || 0))}
                </span>
              </div>
            );
          })}
        </div>

        <div style={{ padding: '10px 16px', borderTop: `1px solid ${LINE}`, display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ flex: 1, fontSize: 12, color: clay.textMuted }}>
            {picked.length ? `${picked.length} selected · folders are shared as links` : 'Tick files or folders — you can pick from several folders and sites.'}
          </div>
          <button type="button" onClick={props.onClose} style={{ padding: '8px 14px', borderRadius: 8, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text, cursor: 'pointer', fontWeight: 600 }}>Cancel</button>
          <button type="button" disabled={!picked.length} onClick={() => { props.onAdd(picked); props.onClose(); }}
            style={{ padding: '8px 16px', borderRadius: 8, border: 'none', background: picked.length ? '#0e7490' : clay.accentSoft, color: '#fff', cursor: picked.length ? 'pointer' : 'not-allowed', fontWeight: 700 }}>
            Add {picked.length || ''}
          </button>
        </div>
      </div>
    </div>
  );
}
