// Ported from the standalone project's SiteToSiteMigration.tsx — a separate
// copy-only flow (never a move) across two SharePoint sites/libraries.
//
// Confirm & Copy starts a background job on the backend; this tab follows it
// over a live NDJSON stream (files/bytes done, speed, time left, per-file
// activity) with pause / resume / cancel, then shows the post-copy
// verification and offers the Excel report. Earlier jobs are listed under
// "Job history" and can be reopened, resumed or retried from there.
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { MigrationApi, errDetail } from './api';
import { pill, primaryBtn, secondaryBtn, tokens, error as errColor, success as okColor, warning as warnColor } from './styles';
import { formatSize } from './fileUtils';
import type {
  S2SConflictPolicy, S2SCopyOptions, S2SDrive, S2SFile, S2SFolder, S2SItem, S2SJob, S2SProgress, S2SProgressEvent, S2SSite,
} from './types';

type Step = 'source' | 'selection' | 'destination' | 'scanning' | 'preview' | 'copying' | 'summary';
type View = 'new' | 'history';

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

const defaultOptions: S2SCopyOptions = { conflictPolicy: 'skip', copyPermissions: false, copyVersions: false };
const MAX_EVENTS = 200;

const CONFLICT_CHOICES: { value: S2SConflictPolicy; label: string; hint: string }[] = [
  { value: 'skip', label: 'Skip it', hint: 'Leave the existing file alone and note it in the report.' },
  { value: 'replace', label: 'Replace it', hint: 'Overwrite the existing file with the source file.' },
  { value: 'rename', label: 'Keep both', hint: 'Copy the new file with a number added, e.g. "Report 1.pdf".' },
  { value: 'fail', label: 'Mark as failed', hint: 'Do not copy it and list it as a failure to review.' },
];

function relOf(fullPath: string, root: string): string {
  if (!root) return fullPath;
  return fullPath === root ? '' : fullPath.startsWith(root + '/') ? fullPath.slice(root.length + 1) : fullPath;
}

function formatDuration(seconds: number | null | undefined): string {
  if (seconds === null || seconds === undefined) return '—';
  if (seconds < 60) return `${Math.max(0, Math.round(seconds))}s`;
  const m = Math.floor(seconds / 60);
  if (m < 60) return `${m}m ${Math.round(seconds % 60)}s`;
  return `${Math.floor(m / 60)}h ${m % 60}m`;
}

function formatWhen(iso: string | null): string {
  if (!iso) return '—';
  const d = new Date(iso.endsWith('Z') || iso.includes('+') ? iso : `${iso}Z`);
  return isNaN(d.getTime()) ? iso : d.toLocaleString();
}

const COPY_STATUS_LABEL: Record<string, [string, 'success' | 'warning' | 'error' | 'primary']> = {
  running: ['Copying', 'primary'],
  paused: ['Paused', 'warning'],
  cancelling: ['Cancelling…', 'warning'],
  verifying: ['Verifying', 'primary'],
  cancelled: ['Cancelled', 'warning'],
  interrupted: ['Interrupted', 'warning'],
  failed: ['Failed', 'error'],
  completed: ['Completed', 'success'],
  completed_with_warnings: ['Completed with warnings', 'warning'],
  completed_with_errors: ['Completed with errors', 'error'],
};

function StatusPill({ status }: { status: string | null | undefined }): React.ReactElement {
  if (!status) return <span style={pill('primary')}>Not copied yet</span>;
  const [label, kind] = COPY_STATUS_LABEL[status] || [status, 'primary'];
  return <span style={pill(kind)}>{label}</span>;
}

export function SiteToSiteTab({ api, isNight }: { api: MigrationApi; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  const [view, setView] = React.useState<View>('new');
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
  const [options, setOptions] = React.useState<S2SCopyOptions>(defaultOptions);
  const [confirming, setConfirming] = React.useState(false);
  const [live, setLive] = React.useState<S2SProgress | null>(null);
  const [events, setEvents] = React.useState<S2SProgressEvent[]>([]);
  const [reconnecting, setReconnecting] = React.useState(false);
  const [controlBusy, setControlBusy] = React.useState(false);
  const [confirmCancel, setConfirmCancel] = React.useState(false);
  const [verifying, setVerifying] = React.useState(false);
  const [history, setHistory] = React.useState<S2SJob[]>([]);
  const [historyLoading, setHistoryLoading] = React.useState(false);

  // Long-running loops (scan polling, the progress stream) check this so
  // they stop when the tab unmounts or the user moves on to another job.
  const mountedRef = React.useRef(true);
  const streamAbortRef = React.useRef<AbortController | null>(null);
  const followedJobRef = React.useRef<string | null>(null);

  React.useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      streamAbortRef.current?.abort();
    };
  }, []);

  React.useEffect(() => {
    setSitesLoading(true);
    api.listS2SSites().then(setSites).catch((e) => setErrorMsg(errDetail(e, 'Could not load configured sites.'))).finally(() => setSitesLoading(false));
  }, [api]);

  const stopFollowing = (): void => {
    followedJobRef.current = null;
    streamAbortRef.current?.abort();
    streamAbortRef.current = null;
  };

  const loadItems = async (jobId: string): Promise<void> => {
    try {
      const data = await api.getS2SJobItems(jobId);
      if (!mountedRef.current) return;
      setJob(data.job);
      setItems(data.items);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load this job’s items.'));
    }
  };

  /** Follow a job's live progress until it finishes, reconnecting if the
   *  stream drops (e.g. a proxy timeout or a brief network blip). */
  const followJob = async (jobId: string): Promise<void> => {
    stopFollowing();
    followedJobRef.current = jobId;
    setEvents([]);
    let failures = 0;
    while (mountedRef.current && followedJobRef.current === jobId) {
      const ctrl = new AbortController();
      streamAbortRef.current = ctrl;
      try {
        // Held in an object: TS doesn't track assignments made in callbacks.
        const result: { finished: S2SJob | null } = { finished: null };
        await api.streamS2SJob(jobId, (line) => {
          if (!mountedRef.current || followedJobRef.current !== jobId) return;
          setReconnecting(false);
          failures = 0;
          if (line.type === 'progress') {
            setLive(line);
            if (line.events.length) setEvents((prev) => [...prev, ...line.events].slice(-MAX_EVENTS));
          } else {
            result.finished = line.job;
          }
        }, ctrl.signal);
        if (result.finished && followedJobRef.current === jobId) {
          followedJobRef.current = null;
          setJob(result.finished);
          setStep('summary');
          await loadItems(jobId);
        }
        return;
      } catch (e) {
        if (ctrl.signal.aborted || !mountedRef.current) return;
        failures += 1;
        setReconnecting(true);
        if (failures >= 20) {
          setErrorMsg(errDetail(e, 'Lost the live progress feed. The copy keeps running on the server — reopen it from Job history.'));
          return;
        }
        await new Promise((r) => setTimeout(r, Math.min(1000 * failures, 10000)));
      }
    }
  };

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

  const pollScan = async (jobId: string): Promise<void> => {
    let failures = 0;
    while (mountedRef.current) {
      try {
        const data = await api.getS2SJobItems(jobId);
        if (!mountedRef.current) return;
        failures = 0;
        setJob(data.job);
        setItems(data.items);
        if (data.job.status === 'done') { setStep('preview'); return; }
        if (data.job.status === 'failed') { setErrorMsg(data.job.error ?? 'Scan failed.'); setStep('destination'); return; }
      } catch (e) {
        failures += 1;
        if (failures >= 5) { setErrorMsg(errDetail(e, 'Lost contact with the scan.')); setStep('destination'); return; }
      }
      await new Promise((r) => setTimeout(r, 1200));
    }
  };

  const startScan = async (): Promise<void> => {
    if (!source.site || !source.drive || !dest.site || !dest.drive) return;
    setErrorMsg(null);
    setStep('scanning');
    setOptions(defaultOptions);
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
    setErrorMsg(null);
    try {
      const started = await api.confirmS2SJob(job.id, options);
      setJob(started);
      setLive(null);
      setConfirmCancel(false);
      setStep('copying');
      void followJob(started.id);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not start the copy.'));
    } finally {
      setConfirming(false);
    }
  };

  const control = async (action: 'pause' | 'resume' | 'cancel'): Promise<void> => {
    if (!job) return;
    setControlBusy(true);
    setErrorMsg(null);
    try {
      const updated = action === 'pause' ? await api.pauseS2SJob(job.id)
        : action === 'resume' ? await api.resumeS2SJob(job.id)
        : await api.cancelS2SJob(job.id);
      setJob(updated);
      setConfirmCancel(false);
      if (action === 'pause') setLive((l) => (l ? { ...l, state: 'paused', speed_bps: 0, eta_seconds: null } : l));
      // A resumed interrupted/cancelled job is a brand-new run on the
      // server — reattach the stream so it's followed again.
      if (action === 'resume' && followedJobRef.current !== job.id) {
        setStep('copying');
        void followJob(job.id);
      }
    } catch (e) {
      setErrorMsg(errDetail(e, `Could not ${action} the copy.`));
    } finally {
      setControlBusy(false);
    }
  };

  const runVerify = async (): Promise<void> => {
    if (!job) return;
    setVerifying(true);
    setErrorMsg(null);
    try {
      await api.verifyS2SJob(job.id);
      await loadItems(job.id);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Verification failed.'));
    } finally {
      setVerifying(false);
    }
  };

  const downloadReport = async (jobId: string): Promise<void> => {
    setErrorMsg(null);
    try {
      const res = await fetch(api.s2sReportUrl(jobId), { headers: api.authHeaders() });
      if (!res.ok) throw new Error(`Report download failed (${res.status})`);
      const blob = await res.blob();
      const href = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = href;
      a.download = `site-to-site-job-${jobId}-report.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(href), 1000);
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Could not download the report.');
    }
  };

  /** Copy the remaining / failed items of a finished job again — back to the
   *  options screen with the job's previous choices pre-filled. */
  const retryJob = (): void => {
    if (!job) return;
    setOptions({
      conflictPolicy: job.conflict_policy || 'skip',
      copyPermissions: job.copy_permissions,
      copyVersions: job.copy_versions,
    });
    setStep('preview');
  };

  const loadHistory = async (): Promise<void> => {
    setHistoryLoading(true);
    try {
      const jobs = await api.listS2SJobs();
      if (mountedRef.current) setHistory(jobs);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load job history.'));
    } finally {
      setHistoryLoading(false);
    }
  };

  const openJob = async (j: S2SJob): Promise<void> => {
    stopFollowing();
    setErrorMsg(null);
    setView('new');
    setJob(j);
    setLive(null);
    setEvents([]);
    setConfirmCancel(false);
    if (j.status === 'running') { setStep('scanning'); await pollScan(j.id); return; }
    if (j.status === 'failed') { setErrorMsg(j.error ?? 'This job’s scan failed.'); await loadItems(j.id); setStep('summary'); return; }
    if (j.live || j.copy_status === 'running' || j.copy_status === 'paused') {
      setStep('copying');
      void followJob(j.id);
      return;
    }
    await loadItems(j.id);
    if (!j.copy_status) {
      setOptions(defaultOptions);
      setStep('preview');
    } else {
      setStep('summary');
    }
  };

  const resetAll = (): void => {
    stopFollowing();
    setStep('source'); setSource(emptyPicker); setDest(emptyPicker); setTreeNodes({});
    setExpandedPaths(new Set()); setCheckedFolders(new Set()); setCheckedFiles(new Set());
    setJob(null); setItems([]); setLive(null); setEvents([]); setErrorMsg(null); setOptions(defaultOptions);
  };

  const fileItems = items.filter((i) => i.kind === 'file');
  const folderCount = items.filter((i) => i.kind === 'folder').length;
  const alreadyCopied = fileItems.filter((i) => i.status === 'copied' || i.status === 'metadata_done').length;
  const toCopy = fileItems.length - alreadyCopied;
  const totalBytes = fileItems.reduce((n, i) => n + (i.size || 0), 0);

  const tabBtn = (active: boolean): React.CSSProperties => ({
    border: 'none', background: 'transparent', cursor: 'pointer', fontSize: 13, padding: '6px 10px',
    color: active ? 'var(--clay-accent)' : t.textMuted, fontWeight: active ? 700 : 500,
    borderBottom: active ? '2px solid var(--clay-accent)' : '2px solid transparent',
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 4px' }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: t.text }}>Site-to-Site Migration (copy, not move)</div>
        <div style={{ marginLeft: 'auto', display: 'flex', gap: 4 }}>
          <button style={tabBtn(view === 'new')} onClick={() => setView('new')}>{job && step !== 'source' ? 'Current job' : 'New copy'}</button>
          <button style={tabBtn(view === 'history')} onClick={() => { setView('history'); void loadHistory(); }}>Job history</button>
        </div>
      </div>
      {errorMsg && <div style={{ borderRadius: 10, border: `1px solid ${errColor}55`, background: `${errColor}15`, color: errColor, padding: '8px 14px', fontSize: 13, marginBottom: 10 }}>{errorMsg}</div>}
      <div style={{ flex: 1, overflowY: 'auto' }}>
        {view === 'history' && (
          <HistoryList jobs={history} loading={historyLoading} isNight={isNight} onRefresh={() => void loadHistory()}
            onOpen={(j) => void openJob(j)} onReport={(j) => void downloadReport(j.id)} />
        )}
        {view === 'new' && step === 'source' && (
          <SitePicker title="Source" api={api} picker={source} sites={sites} sitesLoading={sitesLoading} isNight={isNight}
            onPickSite={(s) => pickSite('source', s)} onPickDrive={(d) => source.site && pickDrive('source', source.site, d)}
            onBrowseInto={(f) => browseInto('source', source, f)} onBrowseUp={() => browseUp('source', source)}
            onContinue={confirmSourceRoot} continueLabel="Next: select folders & files" onError={setErrorMsg} />
        )}
        {view === 'new' && step === 'selection' && (
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
        {view === 'new' && step === 'destination' && (
          <div>
            <button style={secondaryBtn(isNight)} onClick={() => setStep('selection')}>Back</button>
            <div style={{ marginTop: 10 }}>
              <SitePicker title="Destination" api={api} picker={dest} sites={sites} sitesLoading={sitesLoading} isNight={isNight}
                onPickSite={(s) => pickSite('dest', s)} onPickDrive={(d) => dest.site && pickDrive('dest', dest.site, d)}
                onBrowseInto={(f) => browseInto('dest', dest, f)} onBrowseUp={() => browseUp('dest', dest)}
                onContinue={startScan} continueLabel="Scan selection" onError={setErrorMsg} />
            </div>
          </div>
        )}
        {view === 'new' && step === 'scanning' && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10, padding: 60 }}>
            <Icon iconName="Sync" style={{ fontSize: 22, color: 'var(--clay-accent)' }} />
            <div style={{ fontSize: 13, color: t.textMuted }}>Discovering the selected folders and files…{job ? ` ${job.processed} found so far.` : ''}</div>
          </div>
        )}
        {view === 'new' && step === 'preview' && job && (
          <div style={{ maxWidth: 680 }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>{alreadyCopied > 0 ? 'Copy the remaining items' : 'Ready to copy'}</div>
            <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>
              {job.source_site_label} / {job.source_folder_path || '(library root)'} → {job.dest_site_label} / {job.dest_folder_path || '(library root)'}
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 10, marginTop: 12 }}>
              <Tile label="Folders" value={folderCount} isNight={isNight} />
              <Tile label={alreadyCopied > 0 ? 'Files still to copy' : 'Files to copy'} value={toCopy} isNight={isNight} />
              <Tile label="Total size" value={formatSize(totalBytes) || '0 B'} isNight={isNight} />
              {alreadyCopied > 0 && <Tile label="Already copied" value={alreadyCopied} isNight={isNight} />}
            </div>
            <OptionsPanel options={options} onChange={setOptions} isNight={isNight} />
            <div style={{ fontSize: 12, color: t.textMuted, marginTop: 10 }}>Nothing is written to SharePoint until you confirm. The source is never changed.</div>
            <div style={{ display: 'flex', gap: 10, marginTop: 14 }}>
              <button style={secondaryBtn(isNight)} onClick={() => (job.copy_status ? setStep('summary') : setStep('destination'))}>Back</button>
              <button style={primaryBtn(confirming)} disabled={confirming} onClick={() => void doConfirm()}>
                <Icon iconName="CheckMark" style={{ fontSize: 12 }} />{confirming ? 'Starting…' : 'Confirm & Copy'}
              </button>
            </div>
          </div>
        )}
        {view === 'new' && step === 'copying' && job && (
          <LiveProgress job={job} live={live} events={events} reconnecting={reconnecting} isNight={isNight}
            busy={controlBusy} confirmCancel={confirmCancel}
            onPause={() => void control('pause')} onResume={() => void control('resume')}
            onCancel={() => (confirmCancel ? void control('cancel') : setConfirmCancel(true))}
            onKeepGoing={() => setConfirmCancel(false)} />
        )}
        {view === 'new' && step === 'summary' && job && (
          <Summary job={job} items={items} isNight={isNight} verifying={verifying}
            onVerify={() => void runVerify()} onReport={() => void downloadReport(job.id)}
            onRetry={retryJob} onResume={() => void control('resume')} onNew={resetAll} />
        )}
      </div>
    </div>
  );
}

function OptionsPanel({ options, onChange, isNight }: { options: S2SCopyOptions; onChange: (o: S2SCopyOptions) => void; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  const label: React.CSSProperties = { fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase', marginTop: 16 };
  return (
    <div>
      <div style={label}>If a file with the same name already exists</div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 8, marginTop: 6 }}>
        {CONFLICT_CHOICES.map((c) => {
          const active = options.conflictPolicy === c.value;
          return (
            <label key={c.value} style={{
              display: 'block', cursor: 'pointer', borderRadius: 10, padding: '8px 10px',
              border: `1px solid ${active ? 'var(--clay-accent)' : t.border}`, background: active ? 'var(--clay-accent-soft)' : 'transparent',
            }}>
              <input type="radio" name="s2s-conflict" checked={active} onChange={() => onChange({ ...options, conflictPolicy: c.value })} style={{ marginRight: 6 }} />
              <span style={{ fontSize: 12, fontWeight: 600, color: t.text }}>{c.label}</span>
              <div style={{ fontSize: 11, color: t.textMuted, marginTop: 3 }}>{c.hint}</div>
            </label>
          );
        })}
      </div>
      <div style={label}>Also copy</div>
      <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', marginTop: 6, fontSize: 12, color: t.text, cursor: 'pointer' }}>
        <input type="checkbox" checked={options.copyVersions} onChange={(e) => onChange({ ...options, copyVersions: e.target.checked })} />
        <span>Version history<div style={{ fontSize: 11, color: t.textMuted }}>All earlier versions of each file, not just the latest. Slower, and uses more storage at the destination.</div></span>
      </label>
      <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', marginTop: 8, fontSize: 12, color: t.text, cursor: 'pointer' }}>
        <input type="checkbox" checked={options.copyPermissions} onChange={(e) => onChange({ ...options, copyPermissions: e.target.checked })} />
        <span>Unique permissions<div style={{ fontSize: 11, color: t.textMuted }}>People and groups given direct access to a file or folder get the same access on the copy (no email is sent). Sharing links and SharePoint groups can’t be copied and are listed in the report.</div></span>
      </label>
    </div>
  );
}

function ProgressBar({ value, isNight, color }: { value: number; isNight: boolean; color?: string }): React.ReactElement {
  const t = tokens(isNight);
  const pct = Math.max(0, Math.min(100, value * 100));
  return (
    <div style={{ height: 10, borderRadius: 999, background: t.surfaceAlt, border: `1px solid ${t.border}`, overflow: 'hidden' }}>
      <div style={{ width: `${pct}%`, height: '100%', background: color || 'var(--clay-accent)', transition: 'width 0.4s ease' }} />
    </div>
  );
}

const EVENT_STYLE: Record<string, [string, string]> = {
  copied: ['CheckMark', okColor],
  created: ['FabricFolder', okColor],
  attention: ['Warning', warnColor],
  skipped: ['Forward', warnColor],
  failed: ['ErrorBadge', errColor],
};

function LiveProgress({ job, live, events, reconnecting, isNight, busy, confirmCancel, onPause, onResume, onCancel, onKeepGoing }: {
  job: S2SJob; live: S2SProgress | null; events: S2SProgressEvent[]; reconnecting: boolean; isNight: boolean;
  busy: boolean; confirmCancel: boolean; onPause: () => void; onResume: () => void; onCancel: () => void; onKeepGoing: () => void;
}): React.ReactElement {
  const t = tokens(isNight);
  const state = live?.state || job.copy_status || 'running';
  const verifyingNow = live?.phase === 'verifying';
  const ratio = !live ? 0
    : verifyingNow ? (live.verify_total ? live.verify_done / live.verify_total : 1)
    : live.bytes_total > 0 ? live.bytes_done / live.bytes_total
    : live.files_total > 0 ? (live.files_done + live.files_skipped + live.files_failed) / live.files_total : 0;
  const paused = state === 'paused';
  const stopping = state === 'cancelling';

  return (
    <div style={{ maxWidth: 760 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>Copy job #{job.id}</div>
        <StatusPill status={state} />
        {reconnecting && <span style={{ fontSize: 11, color: t.textSubtle }}>Reconnecting to live progress…</span>}
      </div>
      <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>
        {job.source_site_label} / {job.source_folder_path || '(library root)'} → {job.dest_site_label} / {job.dest_folder_path || '(library root)'}
      </div>

      <div style={{ marginTop: 14, display: 'flex', justifyContent: 'space-between', fontSize: 12, color: t.textMuted }}>
        <span>
          {!live ? 'Starting…'
            : verifyingNow ? `Verifying copies — ${live.verify_done} of ${live.verify_total} checked`
            : live.phase === 'folders' ? `Creating folders — ${live.folders_done} of ${live.folders_total}`
            : `${formatSize(live.bytes_done) || '0 B'} of ${formatSize(live.bytes_total) || '0 B'}`}
        </span>
        <span>{Math.round(ratio * 100)}%</span>
      </div>
      <div style={{ marginTop: 6 }}><ProgressBar value={ratio} isNight={isNight} color={paused ? warnColor : undefined} /></div>

      {live && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(120px, 1fr))', gap: 10, marginTop: 14 }}>
          <Tile label="Files copied" value={`${live.files_done} / ${live.files_total}`} isNight={isNight} />
          <Tile label="Speed" value={paused ? 'Paused' : live.speed_bps > 0 ? `${formatSize(live.speed_bps)}/s` : '—'} isNight={isNight} />
          <Tile label="Time left" value={paused ? '—' : formatDuration(live.eta_seconds)} isNight={isNight} />
          <Tile label="Elapsed" value={formatDuration(live.elapsed_seconds)} isNight={isNight} />
          <Tile label="Failed" value={live.files_failed + live.folders_failed} isNight={isNight} tone={live.files_failed + live.folders_failed ? 'error' : undefined} />
          <Tile label="Skipped (already there)" value={live.files_skipped} isNight={isNight} tone={live.files_skipped ? 'warning' : undefined} />
        </div>
      )}

      <div style={{ display: 'flex', gap: 10, marginTop: 14, alignItems: 'center' }}>
        {!verifyingNow && !stopping && (paused
          ? <button style={primaryBtn(busy)} disabled={busy} onClick={onResume}><Icon iconName="Play" style={{ fontSize: 12 }} />Resume</button>
          : <button style={secondaryBtn(isNight)} disabled={busy} onClick={onPause}><Icon iconName="Pause" style={{ fontSize: 12 }} />Pause</button>)}
        {!verifyingNow && !stopping && (confirmCancel ? (
          <>
            <span style={{ fontSize: 12, color: t.text }}>Stop after the files in progress? Copied files stay; you can resume later.</span>
            <button style={{ ...secondaryBtn(isNight), color: errColor, borderColor: `${errColor}66` }} disabled={busy} onClick={onCancel}>Yes, cancel</button>
            <button style={secondaryBtn(isNight)} onClick={onKeepGoing}>Keep copying</button>
          </>
        ) : (
          <button style={{ ...secondaryBtn(isNight), color: errColor }} disabled={busy} onClick={onCancel}><Icon iconName="Cancel" style={{ fontSize: 12 }} />Cancel</button>
        ))}
        {stopping && <span style={{ fontSize: 12, color: t.textMuted }}>Finishing the files in progress, then stopping…</span>}
      </div>

      {live && live.in_flight.length > 0 && !paused && (
        <div style={{ marginTop: 14, fontSize: 12, color: t.textMuted }}>
          <div style={{ fontWeight: 700, textTransform: 'uppercase', fontSize: 11 }}>Copying now</div>
          {live.in_flight.map((p) => <div key={p} style={{ marginTop: 3, color: t.text }}><Icon iconName="Sync" style={{ fontSize: 11, marginRight: 6 }} />{p}</div>)}
        </div>
      )}

      <div style={{ marginTop: 14, fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase' }}>Activity</div>
      <div style={{ marginTop: 4, maxHeight: 260, overflow: 'auto', borderRadius: 10, border: `1px solid ${t.border}` }}>
        {events.length === 0 ? <div style={{ padding: 10, fontSize: 12, color: t.textMuted }}>Waiting for the first files…</div>
          : events.slice().reverse().map((e) => {
            const [icon, color] = EVENT_STYLE[e.status] || ['Info', t.textMuted];
            return (
              <div key={e.seq} style={{ display: 'flex', gap: 8, padding: '6px 10px', fontSize: 12, borderBottom: `1px solid ${t.border}`, color: t.text }}>
                <Icon iconName={icon} style={{ color, fontSize: 12, marginTop: 2 }} />
                <span style={{ flex: 1, wordBreak: 'break-all' }}>{e.path}{e.detail && <span style={{ color: t.textMuted }}> — {e.detail}</span>}</span>
              </div>
            );
          })}
      </div>
    </div>
  );
}

function Summary({ job, items, isNight, verifying, onVerify, onReport, onRetry, onResume, onNew }: {
  job: S2SJob; items: S2SItem[]; isNight: boolean; verifying: boolean;
  onVerify: () => void; onReport: () => void; onRetry: () => void; onResume: () => void; onNew: () => void;
}): React.ReactElement {
  const t = tokens(isNight);
  const s: Partial<S2SProgress> & { error?: string } = job.copy_summary || {};
  const v = job.verify_summary;
  const files = items.filter((i) => i.kind === 'file');
  const copied = files.filter((i) => ['copied', 'metadata_done', 'metadata_attention'].includes(i.status)).length;
  const problems = items.filter((i) => i.status === 'failed' || (i.verify_status && !['ok', 'changed_by_sharepoint'].includes(i.verify_status)));
  const attention = items.filter((i) => i.status === 'metadata_attention' || (i.permissions_report || []).some((p) => p.status === 'error'));
  const remaining = files.filter((i) => ['discovered', 'failed', 'metadata_attention'].includes(i.status)).length;
  const resumable = job.copy_status === 'cancelled' || job.copy_status === 'interrupted' || job.copy_status === 'failed';
  const duration = job.copy_started_at && job.copy_finished_at
    ? (new Date(`${job.copy_finished_at}Z`).getTime() - new Date(`${job.copy_started_at}Z`).getTime()) / 1000 : null;

  return (
    <div style={{ maxWidth: 760 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>Copy job #{job.id}</div>
        <StatusPill status={job.copy_status} />
      </div>
      <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>
        {job.source_site_label} / {job.source_folder_path || '(library root)'} → {job.dest_site_label} / {job.dest_folder_path || '(library root)'}
        {job.confirmed_by_email ? ` · by ${job.confirmed_by_email}` : ''}
      </div>
      {s.error && <div style={{ marginTop: 8, fontSize: 12, color: errColor }}>{s.error}</div>}
      {job.copy_status === 'interrupted' && <div style={{ marginTop: 8, fontSize: 12, color: warnColor }}>The server restarted while this copy was running. Resume to copy the rest — nothing already copied is repeated.</div>}

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(120px, 1fr))', gap: 10, marginTop: 12 }}>
        <Tile label="Files copied" value={`${copied} / ${files.length}`} isNight={isNight} />
        <Tile label="Data copied" value={formatSize(s.bytes_done ?? files.reduce((n, i) => n + (['copied', 'metadata_done', 'metadata_attention'].includes(i.status) ? i.size : 0), 0)) || '0 B'} isNight={isNight} />
        <Tile label="Skipped" value={files.filter((i) => i.status === 'skipped').length} isNight={isNight} tone={files.some((i) => i.status === 'skipped') ? 'warning' : undefined} />
        <Tile label="Failed" value={items.filter((i) => i.status === 'failed').length} isNight={isNight} tone={items.some((i) => i.status === 'failed') ? 'error' : undefined} />
        <Tile label="Needs attention" value={attention.length} isNight={isNight} tone={attention.length ? 'warning' : undefined} />
        <Tile label="Duration" value={formatDuration(duration)} isNight={isNight} />
      </div>

      <div style={{ marginTop: 14, borderRadius: 10, border: `1px solid ${t.border}`, padding: 12 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: t.text }}>Verification</span>
          {job.verify_status === 'running' || verifying ? <span style={pill('primary')}>Checking…</span>
            : !v ? <span style={pill('primary')}>Not run</span>
            : v.passed ? <span style={pill('success')}>Passed</span> : <span style={pill('error')}>Issues found</span>}
          <span style={{ marginLeft: 'auto', fontSize: 11, color: t.textSubtle }}>{job.verified_at ? `Checked ${formatWhen(job.verified_at)}` : ''}</span>
        </div>
        {v && (
          <div style={{ fontSize: 12, color: t.text, marginTop: 8, lineHeight: 1.8 }}>
            Source: {v.source_files} files, {v.source_folders} folders, {formatSize(v.source_bytes) || '0 B'}<br />
            Destination (verified): {v.dest_files_verified} files, {formatSize(v.dest_bytes_verified) || '0 B'}<br />
            Matching: {v.ok}{v.changed_by_sharepoint ? ` · Office files updated by SharePoint metadata: ${v.changed_by_sharepoint}` : ''}
            {v.skipped ? ` · Skipped on purpose: ${v.skipped}` : ''}<br />
            {(v.missing || v.size_mismatch || v.hash_mismatch || v.error || v.not_copied) ? (
              <span style={{ color: errColor }}>
                Missing at destination: {v.missing} · Size differs: {v.size_mismatch} · Content differs: {v.hash_mismatch} · Couldn’t check: {v.error} · Not copied: {v.not_copied}
              </span>
            ) : null}
          </div>
        )}
      </div>

      {problems.length > 0 && (
        <>
          <div style={{ marginTop: 14, fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase' }}>Problems ({problems.length})</div>
          <div style={{ marginTop: 4, maxHeight: 220, overflow: 'auto', borderRadius: 10, border: `1px solid ${t.border}` }}>
            {problems.slice(0, 200).map((i) => (
              <div key={i.id} style={{ padding: '6px 10px', fontSize: 12, borderBottom: `1px solid ${t.border}`, color: t.text }}>
                <div style={{ wordBreak: 'break-all' }}>{i.relative_path}</div>
                <div style={{ color: errColor, fontSize: 11 }}>{i.error || i.verify_detail || i.verify_status}</div>
              </div>
            ))}
          </div>
        </>
      )}

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 16 }}>
        {resumable && <button style={primaryBtn()} onClick={onResume}><Icon iconName="Play" style={{ fontSize: 12 }} />Resume copy</button>}
        {!resumable && remaining > 0 && <button style={primaryBtn()} onClick={onRetry}><Icon iconName="Refresh" style={{ fontSize: 12 }} />Retry {remaining} item(s)</button>}
        <button style={secondaryBtn(isNight)} onClick={onReport}><Icon iconName="ExcelDocument" style={{ fontSize: 12 }} />Download report</button>
        <button style={secondaryBtn(isNight)} disabled={verifying || job.verify_status === 'running'} onClick={onVerify}><Icon iconName="CheckList" style={{ fontSize: 12 }} />Verify again</button>
        <button style={secondaryBtn(isNight)} onClick={onNew}>Start another copy</button>
      </div>
    </div>
  );
}

function HistoryList({ jobs, loading, isNight, onRefresh, onOpen, onReport }: {
  jobs: S2SJob[]; loading: boolean; isNight: boolean; onRefresh: () => void; onOpen: (j: S2SJob) => void; onReport: (j: S2SJob) => void;
}): React.ReactElement {
  const t = tokens(isNight);
  return (
    <div style={{ maxWidth: 900 }}>
      <div style={{ display: 'flex', alignItems: 'center' }}>
        <span style={{ fontSize: 12, color: t.textMuted }}>Every site-to-site copy, newest first. Open one to follow it live, resume it, or see its verification.</span>
        <button style={{ ...secondaryBtn(isNight), marginLeft: 'auto' }} onClick={onRefresh}><Icon iconName="Refresh" style={{ fontSize: 12 }} />Refresh</button>
      </div>
      <div style={{ marginTop: 10, borderRadius: 10, border: `1px solid ${t.border}` }}>
        {loading && jobs.length === 0 ? <div style={{ padding: 12, fontSize: 12, color: t.textMuted }}>Loading…</div>
          : jobs.length === 0 ? <div style={{ padding: 12, fontSize: 12, color: t.textMuted }}>No copy jobs yet.</div>
          : jobs.map((j) => {
            const s = j.copy_summary;
            const scanState = j.status === 'running' ? 'Scanning…' : j.status === 'failed' ? 'Scan failed' : null;
            return (
              <div key={j.id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 12px', borderBottom: `1px solid ${t.border}` }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: t.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    #{j.id} · {j.source_site_label} / {j.source_folder_path || 'root'} → {j.dest_site_label} / {j.dest_folder_path || 'root'}
                  </div>
                  <div style={{ fontSize: 11, color: t.textMuted, marginTop: 2 }}>
                    {formatWhen(j.copy_started_at || j.started_at)}
                    {s && s.files_total !== undefined ? ` · ${s.files_done ?? 0}/${s.files_total} files · ${formatSize(s.bytes_done ?? 0) || '0 B'}` : ` · ${j.total_found} item(s) found`}
                    {s && s.files_failed ? ` · ${s.files_failed} failed` : ''}
                    {j.confirmed_by_email ? ` · ${j.confirmed_by_email}` : ''}
                  </div>
                </div>
                {scanState ? <span style={pill(j.status === 'failed' ? 'error' : 'primary')}>{scanState}</span> : <StatusPill status={j.live ? 'running' : j.copy_status} />}
                {j.verify_summary && <span style={pill(j.verify_summary.passed ? 'success' : 'error')}>{j.verify_summary.passed ? 'Verified' : 'Verify issues'}</span>}
                <button style={secondaryBtn(isNight)} onClick={() => onOpen(j)}>Open</button>
                {j.copy_status && <button style={secondaryBtn(isNight)} title="Download report" onClick={() => onReport(j)}><Icon iconName="ExcelDocument" style={{ fontSize: 12 }} /></button>}
              </div>
            );
          })}
      </div>
    </div>
  );
}

function Tile({ label, value, isNight, tone }: { label: string; value: number | string; isNight: boolean; tone?: 'warning' | 'error' }): React.ReactElement {
  const t = tokens(isNight);
  const color = tone === 'error' ? errColor : tone === 'warning' ? warnColor : t.text;
  return (
    <div style={{ borderRadius: 10, border: `1px solid ${t.border}`, background: t.surface, padding: 12 }}>
      <div style={{ fontSize: 18, fontWeight: 700, color }}>{value}</div>
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

function SitePicker({ title, api, picker, sites, sitesLoading, isNight, onPickSite, onPickDrive, onBrowseInto, onBrowseUp, onContinue, continueLabel, onError }: {
  title: string; api: MigrationApi; picker: SidePicker; sites: S2SSite[]; sitesLoading: boolean; isNight: boolean;
  onPickSite: (s: S2SSite) => void; onPickDrive: (d: S2SDrive) => void; onBrowseInto: (f: S2SFolder) => void;
  onBrowseUp: () => void; onContinue: () => void; continueLabel: string; onError: (msg: string | null) => void;
}): React.ReactElement {
  const t = tokens(isNight);
  const [query, setQuery] = React.useState('');
  const [found, setFound] = React.useState<S2SSite[] | null>(null);
  const [searching, setSearching] = React.useState(false);

  const chip = (active: boolean): React.CSSProperties => ({
    borderRadius: 8, border: `1px solid ${active ? 'var(--clay-accent)' : t.border}`, background: active ? 'var(--clay-accent-soft)' : 'transparent',
    color: active ? 'var(--clay-accent)' : t.text, padding: '6px 12px', fontSize: 12, cursor: 'pointer',
  });

  // One box for both: a pasted SharePoint URL is looked up directly (and
  // access-checked); anything else is a tenant-wide name search.
  const findSites = async (): Promise<void> => {
    const q = query.trim();
    if (!q) return;
    setSearching(true);
    onError(null);
    try {
      if (/sharepoint\.com/i.test(q)) {
        const site = await api.resolveS2SSite(q);
        setFound([site]);
        onPickSite(site);
      } else {
        setFound(await api.searchS2SSites(q));
      }
    } catch (e) {
      setFound([]);
      onError(errDetail(e, 'Site search failed.'));
    } finally {
      setSearching(false);
    }
  };

  const quickKeys = new Set(sites.map((s) => s.key));
  const extra = (found || []).filter((s) => !quickKeys.has(s.key));
  const selectedIsExtra = picker.site && !quickKeys.has(picker.site.key) && !extra.some((s) => s.key === picker.site?.key);

  return (
    <div style={{ maxWidth: 640 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: t.text }}>{title}</div>
      <div style={{ marginTop: 10, fontSize: 11, fontWeight: 700, color: t.textMuted, textTransform: 'uppercase' }}>Site</div>
      <div style={{ display: 'flex', gap: 6, marginTop: 4 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') void findSites(); }}
          placeholder="Search sites by name, or paste a site URL"
          style={{ flex: 1, borderRadius: 8, border: `1px solid ${t.border}`, background: t.surface, color: t.text, padding: '7px 10px', fontSize: 12 }}
        />
        <button style={secondaryBtn(isNight)} disabled={searching || !query.trim()} onClick={() => void findSites()}>
          <Icon iconName="Search" style={{ fontSize: 12 }} />{searching ? 'Finding…' : 'Find'}
        </button>
      </div>
      {found && extra.length === 0 && !searching && found.length === 0 && (
        <div style={{ fontSize: 12, color: t.textMuted, marginTop: 6 }}>No sites found. If you know the site, paste its URL — the app can only see sites it has been granted.</div>
      )}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 8 }}>
        {sitesLoading ? <span style={{ fontSize: 12, color: t.textMuted }}>Loading…</span> : sites.map((s) => (
          <button key={s.key} style={chip(picker.site?.key === s.key)} onClick={() => onPickSite(s)}>{s.label}</button>
        ))}
        {extra.map((s) => <button key={s.key} style={chip(picker.site?.key === s.key)} onClick={() => onPickSite(s)}>{s.label}</button>)}
        {selectedIsExtra && picker.site && <button style={chip(true)}>{picker.site.label}</button>}
      </div>
      {!sitesLoading && (
        <div style={{ fontSize: 11, color: t.textSubtle, marginTop: 4 }}>
          {sites.length === 0
            ? 'No sites yet — add one in Sites → Site Management, or find one above.'
            : 'Sites listed here come from Sites → Site Management; a site added there appears here automatically.'}
        </div>
      )}
      {picker.site?.url && <div style={{ fontSize: 11, color: t.textSubtle, marginTop: 4 }}>{picker.site.url}</div>}
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
