// Ported from the standalone project's SiteToSiteMigration.tsx — a separate
// copy-only flow (never a move) across two SharePoint sites/libraries.
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { MigrationApi, errDetail } from './api';
import { primaryBtn, secondaryBtn, tokens, error as errColor } from './styles';
import { formatSize } from './fileUtils';
import type { S2SConfirmSummary, S2SDrive, S2SFile, S2SFolder, S2SItem, S2SJob, S2SSite } from './types';

type Step = 'source' | 'selection' | 'destination' | 'scanning' | 'preview' | 'summary';

interface SidePicker {
  site: S2SSite | null;
  drives: S2SDrive[];
  drive: S2SDrive | null;
  path: string;
  folders: S2SFolder[];
  files: S2SFile[];
  loading: boolean;
}
const emptyPicker: SidePicker = { site: null, drives: [], drive: null, path: '', folders: [], files: [], loading: false };

interface TreeNodeState { folders: S2SFolder[]; files: S2SFile[]; loading: boolean }

function relOf(fullPath: string, root: string): string {
  if (!root) return fullPath;
  return fullPath === root ? '' : fullPath.startsWith(root + '/') ? fullPath.slice(root.length + 1) : fullPath;
}

export function SiteToSiteTab({ api, isNight }: { api: MigrationApi; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  const [step, setStep] = React.useState<Step>('source');
  const [errorMsg, setErrorMsg] = React.useState<string | null>(null);
  const [sites, setSites] = React.useState<S2SSite[]>([]);
  const [sitesLoading, setSitesLoading] = React.useState(true);
  const [source, setSource] = React.useState<SidePicker>(emptyPicker);
  const [dest, setDest] = React.useState<SidePicker>(emptyPicker);
  const [treeNodes, setTreeNodes] = React.useState<Record<string, TreeNodeState>>({});
  const [expandedPaths, setExpandedPaths] = React.useState<Set<string>>(new Set());
  const [checkedFolders, setCheckedFolders] = React.useState<Set<string>>(new Set());
  const [checkedFiles, setCheckedFiles] = React.useState<Set<string>>(new Set());
  const [job, setJob] = React.useState<S2SJob | null>(null);
  const [items, setItems] = React.useState<S2SItem[]>([]);
  const [summary, setSummary] = React.useState<S2SConfirmSummary | null>(null);
  const [confirming, setConfirming] = React.useState(false);

  React.useEffect(() => {
    setSitesLoading(true);
    api.listS2SSites().then(setSites).catch((e) => setErrorMsg(errDetail(e, 'Could not load configured sites.'))).finally(() => setSitesLoading(false));
  }, [api]);

  const pickSite = async (side: 'source' | 'dest', site: S2SSite): Promise<void> => {
    const setter = side === 'source' ? setSource : setDest;
    setter({ ...emptyPicker, site, loading: true });
    try {
      const drives = await api.listS2SSiteDrives(site.key);
      setter((p) => ({ ...p, drives, loading: false }));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load libraries for this site.'));
      setter((p) => ({ ...p, loading: false }));
    }
  };

  const pickDrive = async (side: 'source' | 'dest', site: S2SSite, drive: S2SDrive): Promise<void> => {
    const setter = side === 'source' ? setSource : setDest;
    setter((p) => ({ ...p, drive, path: '', loading: true }));
    try {
      const data = await api.browseS2SFolder(site.key, drive.id);
      setter((p) => ({ ...p, folders: data.folders, files: data.files, loading: false }));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not browse this library.'));
      setter((p) => ({ ...p, loading: false }));
    }
  };

  const browseInto = async (side: 'source' | 'dest', picker: SidePicker, folder: S2SFolder): Promise<void> => {
    if (!picker.site || !picker.drive) return;
    const setter = side === 'source' ? setSource : setDest;
    setter((p) => ({ ...p, path: folder.path, loading: true }));
    try {
      const data = await api.browseS2SFolder(picker.site.key, picker.drive.id, folder.path);
      setter((p) => ({ ...p, folders: data.folders, files: data.files, loading: false }));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not browse this folder.'));
      setter((p) => ({ ...p, loading: false }));
    }
  };

  const browseUp = async (side: 'source' | 'dest', picker: SidePicker): Promise<void> => {
    if (!picker.site || !picker.drive) return;
    const parent = picker.path.split('/').slice(0, -1).join('/');
    const setter = side === 'source' ? setSource : setDest;
    setter((p) => ({ ...p, path: parent, loading: true }));
    try {
      const data = await api.browseS2SFolder(picker.site.key, picker.drive.id, parent || undefined);
      setter((p) => ({ ...p, folders: data.folders, files: data.files, loading: false }));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not browse this folder.'));
      setter((p) => ({ ...p, loading: false }));
    }
  };

  const confirmSourceRoot = (): void => {
    setTreeNodes({ [source.path]: { folders: source.folders, files: source.files, loading: false } });
    setExpandedPaths(new Set([source.path]));
    setCheckedFolders(new Set());
    setCheckedFiles(new Set());
    setStep('selection');
  };

  const toggleExpand = async (path: string): Promise<void> => {
    if (expandedPaths.has(path)) {
      setExpandedPaths((prev) => { const n = new Set(prev); n.delete(path); return n; });
      return;
    }
    setExpandedPaths((prev) => new Set(prev).add(path));
    if (treeNodes[path] || !source.site || !source.drive) return;
    setTreeNodes((prev) => ({ ...prev, [path]: { folders: [], files: [], loading: true } }));
    try {
      const data = await api.browseS2SFolder(source.site.key, source.drive.id, path);
      setTreeNodes((prev) => ({ ...prev, [path]: { folders: data.folders, files: data.files, loading: false } }));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not browse this folder.'));
      setTreeNodes((prev) => ({ ...prev, [path]: { folders: [], files: [], loading: false } }));
    }
  };

  const toggleFolderCheck = (path: string): void => setCheckedFolders((prev) => { const n = new Set(prev); if (n.has(path)) n.delete(path); else n.add(path); return n; });
  const toggleFileCheck = (path: string): void => setCheckedFiles((prev) => { const n = new Set(prev); if (n.has(path)) n.delete(path); else n.add(path); return n; });

  const pollScan = React.useCallback(async (jobId: string) => {
    for (;;) {
      const data = await api.getS2SJobItems(jobId);
      setJob(data.job);
      setItems(data.items);
      if (data.job.status === 'done') { setStep('preview'); return; }
      if (data.job.status === 'failed') { setErrorMsg(data.job.error ?? 'Scan failed.'); setStep('destination'); return; }
      await new Promise((r) => setTimeout(r, 1200));
    }
  }, [api]);

  const startCopy = async (): Promise<void> => {
    if (!source.site || !source.drive || !dest.site || !dest.drive) return;
    setErrorMsg(null);
    setStep('scanning');
    try {
      const newJob = await api.startS2SScan({
        sourceSiteKey: source.site.key, sourceDriveId: source.drive.id, sourceFolderPath: source.path,
        selectedFolders: Array.from(checkedFolders).map((p) => relOf(p, source.path)),
        selectedFiles: Array.from(checkedFiles).map((p) => relOf(p, source.path)),
        destSiteKey: dest.site.key, destDriveId: dest.drive.id, destFolderPath: dest.path,
      });
      setJob(newJob);
      await pollScan(newJob.id);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not start the scan.'));
      setStep('destination');
    }
  };

  const doConfirm = async (): Promise<void> => {
    if (!job) return;
    setConfirming(true);
    try {
      setSummary(await api.confirmS2SJob(job.id));
      setStep('summary');
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not complete the copy.'));
    } finally {
      setConfirming(false);
    }
  };

  const resetAll = (): void => {
    setStep('source'); setSource(emptyPicker); setDest(emptyPicker); setTreeNodes({});
    setExpandedPaths(new Set()); setCheckedFolders(new Set()); setCheckedFiles(new Set());
    setJob(null); setItems([]); setSummary(null); setErrorMsg(null);
  };

  const folderCount = items.filter((i) => i.kind === 'folder').length;
  const fileCount = items.filter((i) => i.kind === 'file').length;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ padding: '14px 4px', fontSize: 15, fontWeight: 700, color: t.text }}>Site-to-Site Migration (copy, not move)</div>
      {errorMsg && <div style={{ borderRadius: 10, border: `1px solid ${errColor}55`, background: `${errColor}15`, color: errColor, padding: '8px 14px', fontSize: 13, marginBottom: 10 }}>{errorMsg}</div>}
      <div style={{ flex: 1, overflowY: 'auto' }}>
        {step === 'source' && (
          <SitePicker title="Source" picker={source} sites={sites} sitesLoading={sitesLoading} isNight={isNight}
            onPickSite={(s) => pickSite('source', s)} onPickDrive={(d) => source.site && pickDrive('source', source.site, d)}
            onBrowseInto={(f) => browseInto('source', source, f)} onBrowseUp={() => browseUp('source', source)}
            onContinue={confirmSourceRoot} continueLabel="Next: select folders & files" />
        )}
        {step === 'selection' && (
          <div style={{ maxWidth: 640 }}>
            <button style={secondaryBtn(isNight)} onClick={() => setStep('source')}>Back</button>
            <div style={{ fontSize: 14, fontWeight: 700, color: t.text, marginTop: 10 }}>Select what to copy</div>
            <div style={{ fontSize: 12, color: t.textMuted, marginTop: 4 }}>Under {source.path || '(root)'}: check a folder to include it fully, or expand and check individual files.</div>
            <div style={{ marginTop: 10, maxHeight: 380, overflow: 'auto', borderRadius: 10, border: `1px solid ${t.border}` }}>
              <TreeLevel path={source.path} depth={0} isNight={isNight} treeNodes={treeNodes} expandedPaths={expandedPaths}
                checkedFolders={checkedFolders} checkedFiles={checkedFiles} onToggleExpand={toggleExpand}
                onToggleFolderCheck={toggleFolderCheck} onToggleFileCheck={toggleFileCheck} />
            </div>
            <div style={{ marginTop: 8, fontSize: 12, color: t.textMuted }}>{checkedFolders.size} folder(s), {checkedFiles.size} file(s) selected</div>
            <button style={{ ...primaryBtn(checkedFolders.size === 0 && checkedFiles.size === 0), marginTop: 10 }} disabled={checkedFolders.size === 0 && checkedFiles.size === 0}
              onClick={() => setStep('destination')}>Next: pick destination</button>
          </div>
        )}
        {step === 'destination' && (
          <div>
            <button style={secondaryBtn(isNight)} onClick={() => setStep('selection')}>Back</button>
            <div style={{ marginTop: 10 }}>
              <SitePicker title="Destination" picker={dest} sites={sites} sitesLoading={sitesLoading} isNight={isNight}
                onPickSite={(s) => pickSite('dest', s)} onPickDrive={(d) => dest.site && pickDrive('dest', dest.site, d)}
                onBrowseInto={(f) => browseInto('dest', dest, f)} onBrowseUp={() => browseUp('dest', dest)}
                onContinue={startCopy} continueLabel="Scan selection" />
            </div>
          </div>
        )}
        {step === 'scanning' && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10, padding: 60 }}>
            <Icon iconName="Sync" style={{ fontSize: 22, color: '#DD9159' }} />
            <div style={{ fontSize: 13, color: t.textMuted }}>Discovering the selected folders and files…{job ? ` ${job.processed} found so far.` : ''}</div>
          </div>
        )}
        {step === 'preview' && job && (
          <div style={{ maxWidth: 680 }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>Ready to copy</div>
            <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>Nothing has been written to SharePoint yet.</div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginTop: 12 }}>
              <Tile label="Folders to create" value={folderCount} isNight={isNight} />
              <Tile label="Files to copy" value={fileCount} isNight={isNight} />
            </div>
            <div style={{ display: 'flex', gap: 10, marginTop: 14 }}>
              <button style={secondaryBtn(isNight)} onClick={() => setStep('destination')}>Back</button>
              <button style={primaryBtn(confirming)} disabled={confirming} onClick={doConfirm}><Icon iconName="CheckMark" style={{ fontSize: 12 }} />Confirm & Copy</button>
            </div>
          </div>
        )}
        {step === 'summary' && summary && (
          <div style={{ maxWidth: 560 }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>Migration Complete</div>
            <div style={{ marginTop: 10, fontSize: 12, color: t.text, lineHeight: 1.9 }}>
              Folders — created: {summary.folders_created}, existing: {summary.folders_existing}, failed: {summary.folders_failed}<br />
              Files — copied: {summary.files_copied}, failed: {summary.files_failed}<br />
              Metadata — migrated: {summary.metadata_migrated}, attention: {summary.metadata_attention}<br />
              Managed Metadata — applied: {summary.managed_metadata_applied}, unmapped: {summary.managed_metadata_unmapped}
            </div>
            <button style={{ ...primaryBtn(), marginTop: 14 }} onClick={resetAll}>Start another copy</button>
          </div>
        )}
      </div>
    </div>
  );
}

function Tile({ label, value, isNight }: { label: string; value: number; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  return (
    <div style={{ borderRadius: 10, border: `1px solid ${t.border}`, background: t.surface, padding: 12 }}>
      <div style={{ fontSize: 20, fontWeight: 700, color: t.text }}>{value}</div>
      <div style={{ fontSize: 11, color: t.textMuted }}>{label}</div>
    </div>
  );
}

function TreeLevel({ path, depth, isNight, treeNodes, expandedPaths, checkedFolders, checkedFiles, onToggleExpand, onToggleFolderCheck, onToggleFileCheck }: {
  path: string; depth: number; isNight: boolean; treeNodes: Record<string, TreeNodeState>; expandedPaths: Set<string>;
  checkedFolders: Set<string>; checkedFiles: Set<string>; onToggleExpand: (p: string) => void;
  onToggleFolderCheck: (p: string) => void; onToggleFileCheck: (p: string) => void;
}): React.ReactElement | null {
  const t = tokens(isNight);
  const node = treeNodes[path];
  if (!node) return null;
  if (node.loading) return <div style={{ padding: 8, paddingLeft: depth * 18 + 10, fontSize: 12, color: t.textMuted }}>Loading…</div>;
  if (node.folders.length === 0 && node.files.length === 0) return <div style={{ padding: 8, paddingLeft: depth * 18 + 10, fontSize: 12, color: t.textMuted }}>Empty.</div>;
  return (
    <>
      {node.folders.map((f) => {
        const isExpanded = expandedPaths.has(f.path);
        return (
          <div key={f.id}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 8px', paddingLeft: depth * 18 + 10, borderBottom: `1px solid ${t.border}` }}>
              <button onClick={() => onToggleExpand(f.path)} style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: t.textMuted }}>
                <Icon iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 11 }} />
              </button>
              <input type="checkbox" checked={checkedFolders.has(f.path)} onChange={() => onToggleFolderCheck(f.path)} />
              <Icon iconName="FabricFolder" style={{ color: t.textMuted, fontSize: 13 }} />
              <span style={{ fontSize: 12, color: t.text }}>{f.name}</span>
            </div>
            {isExpanded && <TreeLevel path={f.path} depth={depth + 1} isNight={isNight} treeNodes={treeNodes} expandedPaths={expandedPaths}
              checkedFolders={checkedFolders} checkedFiles={checkedFiles} onToggleExpand={onToggleExpand}
              onToggleFolderCheck={onToggleFolderCheck} onToggleFileCheck={onToggleFileCheck} />}
          </div>
        );
      })}
      {node.files.map((file) => {
        const filePath = path ? `${path}/${file.name}` : file.name;
        return (
          <div key={file.id} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 8px', paddingLeft: depth * 18 + 30, borderBottom: `1px solid ${t.border}` }}>
            <input type="checkbox" checked={checkedFiles.has(filePath)} onChange={() => onToggleFileCheck(filePath)} />
            <Icon iconName="Page" style={{ color: t.textMuted, fontSize: 13 }} />
            <span style={{ flex: 1, fontSize: 12, color: t.text }}>{file.name}</span>
            <span style={{ fontSize: 11, color: t.textSubtle }}>{formatSize(file.size)}</span>
          </div>
        );
      })}
    </>
  );
}

function SitePicker({ title, picker, sites, sitesLoading, isNight, onPickSite, onPickDrive, onBrowseInto, onBrowseUp, onContinue, continueLabel }: {
  title: string; picker: SidePicker; sites: S2SSite[]; sitesLoading: boolean; isNight: boolean;
  onPickSite: (s: S2SSite) => void; onPickDrive: (d: S2SDrive) => void; onBrowseInto: (f: S2SFolder) => void;
  onBrowseUp: () => void; onContinue: () => void; continueLabel: string;
}): React.ReactElement {
  const t = tokens(isNight);
  const chip = (active: boolean): React.CSSProperties => ({
    borderRadius: 8, border: `1px solid ${active ? '#DD9159' : t.border}`, background: active ? '#DD915922' : 'transparent',
    color: active ? '#DD9159' : t.text, padding: '6px 12px', fontSize: 12, cursor: 'pointer',
  });
  return (
    <div style={{ maxWidth: 640 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>{title}</div>
      <div style={{ marginTop: 10, fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase' }}>Site</div>
      {sitesLoading ? <div style={{ fontSize: 12, color: t.textMuted }}>Loading…</div> : sites.length === 0 ? (
        <div style={{ fontSize: 12, color: t.textMuted }}>No sites configured — set ALLOWED_SITES in the migration backend's .env.</div>
      ) : (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 4 }}>
          {sites.map((s) => <button key={s.key} style={chip(picker.site?.key === s.key)} onClick={() => onPickSite(s)}>{s.label}</button>)}
        </div>
      )}
      {picker.site && (
        <>
          <div style={{ marginTop: 12, fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase' }}>Library</div>
          {picker.loading && !picker.drive ? <div style={{ fontSize: 12, color: t.textMuted }}>Loading…</div> : picker.drives.length === 0 ? (
            <div style={{ fontSize: 12, color: t.textMuted }}>No document library found for this site.</div>
          ) : (
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 4 }}>
              {picker.drives.map((d) => <button key={d.id} style={chip(picker.drive?.id === d.id)} onClick={() => onPickDrive(d)}>{d.name}</button>)}
            </div>
          )}
        </>
      )}
      {picker.drive && (
        <>
          <div style={{ marginTop: 12, display: 'flex', justifyContent: 'space-between' }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase' }}>Folder</span>
            <span style={{ fontSize: 11, color: t.textSubtle }}>/{picker.path}</span>
          </div>
          <div style={{ marginTop: 4, borderRadius: 10, border: `1px solid ${t.border}` }}>
            {picker.path && <div onClick={onBrowseUp} style={{ padding: '8px 10px', fontSize: 12, color: t.textMuted, cursor: 'pointer', borderBottom: `1px solid ${t.border}` }}>.. (up)</div>}
            {picker.loading ? <div style={{ padding: 10, fontSize: 12 }}>Loading…</div> : picker.folders.length === 0 && picker.files.length === 0 ? (
              <div style={{ padding: 10, fontSize: 12, color: t.textMuted }}>Empty.</div>
            ) : (
              <>
                {picker.folders.map((f) => (
                  <div key={f.id} onClick={() => onBrowseInto(f)} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 10px', fontSize: 12, cursor: 'pointer', borderBottom: `1px solid ${t.border}`, color: t.text }}>
                    <span style={{ display: 'flex', gap: 6, alignItems: 'center' }}><Icon iconName="FabricFolder" style={{ fontSize: 12, color: t.textMuted }} />{f.name}</span>
                    <Icon iconName="ChevronRight" style={{ fontSize: 11, color: t.textMuted }} />
                  </div>
                ))}
                {picker.files.map((file) => (
                  <div key={file.id} style={{ display: 'flex', gap: 6, alignItems: 'center', padding: '8px 10px', fontSize: 12, color: t.textMuted, borderBottom: `1px solid ${t.border}` }}>
                    <Icon iconName="Page" style={{ fontSize: 12 }} />{file.name}<span style={{ marginLeft: 'auto', fontSize: 11 }}>{formatSize(file.size)}</span>
                  </div>
                ))}
              </>
            )}
          </div>
          <button style={{ ...primaryBtn(), marginTop: 12 }} onClick={onContinue}>{continueLabel}</button>
        </>
      )}
    </div>
  );
}
