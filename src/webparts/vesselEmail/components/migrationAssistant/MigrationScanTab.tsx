// The scan -> subfolders -> vessel -> preview -> confirm workflow. Ported
// from the standalone SharePoint AI Migration Assistant's
// frontend/src/components/MigrationAssistant.tsx, restyled with inline
// styles (no Tailwind here) to sit natively inside the SPFx web part.
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { MigrationApi, errDetail } from './api';
import { fileIconName, formatSize } from './fileUtils';
import { pill, primaryBtn, secondaryBtn, tokens, error as errColor, success as successColor, warning as warnColor } from './styles';
import type { ConfirmSummary, MigrationItem, MigrationScanJob, SourceFile, SourceFolder } from './types';

type Step = 'source' | 'subfolders' | 'vessel' | 'scanning' | 'preview' | 'summary';

const STEP_LABELS: { key: Step; label: string }[] = [
  { key: 'source', label: '1. Source folder' },
  { key: 'subfolders', label: '2. Subfolders' },
  { key: 'vessel', label: '3. Destination vessel' },
  { key: 'preview', label: '4. Review & confirm' },
];

function confidenceLabel(item: MigrationItem): { text: string; color: string } {
  if (item.overridden) return { text: 'Manually set', color: '#6366f1' };
  if (item.confidence == null) return { text: '—', color: '#94a3b8' };
  const pct = Math.round(item.confidence * 100);
  const color = pct >= 75 ? successColor : pct >= 40 ? warnColor : errColor;
  return { text: `${pct}%`, color };
}

export function MigrationScanTab({ api, isNight, actingEmail }: { api: MigrationApi; isNight: boolean; actingEmail: string }): React.ReactElement {
  const t = tokens(isNight);
  const [step, setStep] = React.useState<Step>('source');
  const [errorMsg, setErrorMsg] = React.useState<string | null>(null);

  const [browsePath, setBrowsePath] = React.useState('');
  const [browseFolders, setBrowseFolders] = React.useState<SourceFolder[]>([]);
  const [browseFiles, setBrowseFiles] = React.useState<SourceFile[]>([]);
  const [browseLoading, setBrowseLoading] = React.useState(true);
  const [sourceFolder, setSourceFolder] = React.useState<SourceFolder | null>(null);

  const [subfolderOptions, setSubfolderOptions] = React.useState<SourceFolder[]>([]);
  const [subfolderFiles, setSubfolderFiles] = React.useState<SourceFile[]>([]);
  const [subfolderLoading, setSubfolderLoading] = React.useState(false);
  const [checkedSubfolders, setCheckedSubfolders] = React.useState<Set<string>>(new Set());
  const [checkedFiles, setCheckedFiles] = React.useState<Set<string>>(new Set());

  const [vessels, setVessels] = React.useState<SourceFolder[]>([]);
  const [destinationRoot, setDestinationRoot] = React.useState('');
  const [vesselLoading, setVesselLoading] = React.useState(false);

  const [job, setJob] = React.useState<MigrationScanJob | null>(null);
  const [items, setItems] = React.useState<MigrationItem[]>([]);
  const [hierarchyPaths, setHierarchyPaths] = React.useState<string[]>([]);
  const [busyItemId, setBusyItemId] = React.useState<string | null>(null);
  const [previewItem, setPreviewItem] = React.useState<MigrationItem | null>(null);

  const [summary, setSummary] = React.useState<ConfirmSummary | null>(null);
  const [confirming, setConfirming] = React.useState(false);

  const [recentOpen, setRecentOpen] = React.useState(false);
  const [recentJobs, setRecentJobs] = React.useState<MigrationScanJob[]>([]);
  const [recentLoading, setRecentLoading] = React.useState(false);

  const loadBrowse = React.useCallback(
    async (path: string) => {
      setBrowseLoading(true);
      setErrorMsg(null);
      try {
        const data = await api.listSourceFolders(path || undefined);
        setBrowsePath(data.path);
        setBrowseFolders(data.folders);
        setBrowseFiles(data.files);
      } catch (e) {
        setErrorMsg(errDetail(e, 'Could not load folders.'));
      } finally {
        setBrowseLoading(false);
      }
    },
    [api]
  );

  React.useEffect(() => {
    if (step === 'source') void loadBrowse(browsePath);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [step, browsePath]);

  const crumbs = browsePath.split('/').filter(Boolean);

  const chooseSourceFolder = async (folder: SourceFolder): Promise<void> => {
    setSourceFolder(folder);
    setCheckedSubfolders(new Set());
    setCheckedFiles(new Set());
    setSubfolderLoading(true);
    setErrorMsg(null);
    setStep('subfolders');
    try {
      const data = await api.listSourceFolders(folder.path);
      setSubfolderOptions(data.folders);
      setSubfolderFiles(data.files);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load subfolders.'));
    } finally {
      setSubfolderLoading(false);
    }
  };

  const toggleSubfolder = (name: string): void =>
    setCheckedSubfolders((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name); else next.add(name);
      return next;
    });

  const toggleFile = (name: string): void =>
    setCheckedFiles((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name); else next.add(name);
      return next;
    });

  const goToVesselStep = async (): Promise<void> => {
    setStep('vessel');
    setVesselLoading(true);
    setErrorMsg(null);
    try {
      const data = await api.listVessels();
      setVessels(data.vessels);
      setDestinationRoot(data.destination_root);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load vessel folders.'));
    } finally {
      setVesselLoading(false);
    }
  };

  const pollJob = React.useCallback(
    (jobId: string) => {
      let failures = 0;
      const interval = window.setInterval(async () => {
        try {
          const polled = await api.getMigrationScanJob(jobId);
          failures = 0;
          setJob(polled);
          if (polled.status !== 'running') {
            window.clearInterval(interval);
            if (polled.status === 'failed') {
              setErrorMsg(polled.error || 'Scan failed.');
              setStep('vessel');
            } else {
              try {
                const [itemsData, hierarchyData] = await Promise.all([api.getJobItems(jobId), api.getJobHierarchy(jobId)]);
                setJob(itemsData.job);
                setItems(itemsData.items);
                setHierarchyPaths(hierarchyData.paths);
                setStep('preview');
              } catch (e) {
                setErrorMsg(errDetail(e, 'Could not load the scan results.'));
              }
            }
          }
        } catch (e) {
          failures += 1;
          if (failures >= 5) {
            window.clearInterval(interval);
            setErrorMsg(`${errDetail(e, 'Could not reach the scan service.')} The scan may still be running — reopen it from Recent scans.`);
          }
        }
      }, 2000);
    },
    [api]
  );

  const startScan = async (vessel: SourceFolder): Promise<void> => {
    if (!sourceFolder) return;
    setStep('scanning');
    setErrorMsg(null);
    try {
      const newJob = await api.startMigrationScan(sourceFolder.path, Array.from(checkedSubfolders), vessel.path, Array.from(checkedFiles));
      setJob(newJob);
      pollJob(newJob.id);
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not start the scan.'));
      setStep('vessel');
    }
  };

  const handleOverride = async (item: MigrationItem, targetPath: string): Promise<void> => {
    if (!job || !targetPath) return;
    setBusyItemId(item.id);
    try {
      const updated = await api.overrideJobItem(job.id, item.id, targetPath);
      setItems((prev) => prev.map((i) => (i.id === updated.id ? updated : i)));
    } catch (e) {
      setErrorMsg(errDetail(e, "Could not update this item's category."));
    } finally {
      setBusyItemId(null);
    }
  };

  const handleReclassify = async (item: MigrationItem): Promise<void> => {
    if (!job) return;
    setBusyItemId(item.id);
    try {
      const updated = await api.reclassifyJobItem(job.id, item.id);
      setItems((prev) => prev.map((i) => (i.id === updated.id ? updated : i)));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not re-run classification.'));
    } finally {
      setBusyItemId(null);
    }
  };

  const handleConfirm = async (): Promise<void> => {
    if (!job) return;
    const destinationLabel = job.auto_detect_vessel ? `their detected vessels under "${job.vessel_name}"` : `"${job.vessel_name}"`;
    // eslint-disable-next-line no-alert
    if (!window.confirm(`Move ${items.length} file(s) into ${destinationLabel} now? Every file must have a valid Category tag. This cannot be undone from here.`)) return;
    setConfirming(true);
    setErrorMsg(null);
    try {
      const result = await api.confirmJob(job.id);
      setSummary(result);
      setStep('summary');
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not confirm the move.'));
    } finally {
      setConfirming(false);
    }
  };

  const startOver = (): void => {
    setStep('source');
    setSourceFolder(null);
    setCheckedSubfolders(new Set());
    setCheckedFiles(new Set());
    setJob(null);
    setItems([]);
    setSummary(null);
    setErrorMsg(null);
    setBrowsePath('');
  };

  const openRecentJobs = async (): Promise<void> => {
    setRecentOpen(true);
    setRecentLoading(true);
    setErrorMsg(null);
    try {
      setRecentJobs(await api.listMigrationJobs());
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load recent scans.'));
    } finally {
      setRecentLoading(false);
    }
  };

  const viewJob = async (target: MigrationScanJob): Promise<void> => {
    setRecentOpen(false);
    setErrorMsg(null);
    if (target.status === 'running') {
      setJob(target);
      setStep('scanning');
      pollJob(target.id);
      return;
    }
    try {
      const [itemsData, hierarchyData] = await Promise.all([api.getJobItems(target.id), api.getJobHierarchy(target.id)]);
      setJob(itemsData.job);
      setItems(itemsData.items);
      setHierarchyPaths(hierarchyData.paths);
      setStep('preview');
    } catch (e) {
      setErrorMsg(errDetail(e, "Could not load that scan's results."));
    }
  };

  const rowStyle: React.CSSProperties = { display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', borderBottom: `1px solid ${t.border}` };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 4px' }}>
        <div>
          <div style={{ fontSize: 15, fontWeight: 700, color: t.text }}>Pick a source folder, then subfolders, then a vessel</div>
          <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>Nothing moves in SharePoint until you click "Confirm Move" on the review table.</div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button style={secondaryBtn(isNight)} onClick={openRecentJobs}><Icon iconName="History" style={{ fontSize: 13 }} />Recent scans</button>
          {step !== 'source' && step !== 'summary' && (
            <button style={secondaryBtn(isNight)} onClick={startOver}><Icon iconName="NavigateBack" style={{ fontSize: 13 }} />Start over</button>
          )}
        </div>
      </div>

      {step !== 'summary' && (
        <div style={{ display: 'flex', gap: 4, padding: '4px 4px 12px', fontSize: 12 }}>
          {STEP_LABELS.map((s, i) => {
            const order = STEP_LABELS.map((x) => x.key);
            const currentIdx = order.indexOf(step === 'scanning' ? 'preview' : step);
            const isDone = i < currentIdx;
            const isCurrent = i === currentIdx;
            return (
              <span key={s.key} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                {i > 0 && <Icon iconName="ChevronRight" style={{ fontSize: 10, color: t.textSubtle }} />}
                <span style={{
                  borderRadius: 999, padding: '3px 10px', fontWeight: 600,
                  background: isCurrent ? 'var(--clay-accent-soft)' : 'transparent',
                  color: isCurrent ? 'var(--clay-accent)' : isDone ? successColor : t.textSubtle,
                }}>{s.label}</span>
              </span>
            );
          })}
        </div>
      )}

      {errorMsg && (
        <div style={{ borderRadius: 10, border: `1px solid ${errColor}55`, background: `${errColor}15`, color: errColor, padding: '8px 14px', fontSize: 13, marginBottom: 10 }}>
          {errorMsg}
        </div>
      )}

      <div style={{ flex: 1, overflowY: 'auto' }}>
        {step === 'source' && (
          <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface }}>
            <div style={{ padding: '12px 14px', borderBottom: `1px solid ${t.border}` }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: t.text }}>Select a source folder</div>
                {crumbs.length > 0 && (
                  <button style={primaryBtn()} onClick={() => chooseSourceFolder({ id: '', name: crumbs[crumbs.length - 1], path: browsePath })}>
                    Use current folder ({crumbs[crumbs.length - 1]})
                  </button>
                )}
              </div>
              <div style={{ marginTop: 4, fontSize: 11, color: t.textSubtle, display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                <a onClick={() => setBrowsePath('')} style={{ cursor: 'pointer', color: t.textMuted }}>Documents</a>
                {crumbs.map((c, i) => (
                  <span key={i} style={{ display: 'flex', gap: 4 }}>
                    <Icon iconName="ChevronRight" style={{ fontSize: 9 }} />
                    <a onClick={() => setBrowsePath(crumbs.slice(0, i + 1).join('/'))} style={{ cursor: 'pointer' }}>{c}</a>
                  </span>
                ))}
              </div>
            </div>
            {browseLoading ? (
              <div style={{ padding: 30, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>Loading…</div>
            ) : browseFolders.length === 0 && browseFiles.length === 0 ? (
              <div style={{ padding: 30, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>This folder is empty.</div>
            ) : (
              <>
                {browseFolders.map((f) => (
                  <div key={f.id} style={rowStyle}>
                    <Icon iconName="FabricFolder" style={{ color: t.textSubtle }} />
                    <span style={{ flex: 1, color: t.text, fontSize: 13, cursor: 'pointer' }} onClick={() => setBrowsePath(f.path)}>{f.name}</span>
                    <button style={primaryBtn()} onClick={() => chooseSourceFolder(f)}>Use this folder</button>
                  </div>
                ))}
                {browseFiles.map((f) => <FileRow key={f.id} file={f} isNight={isNight} />)}
              </>
            )}
          </div>
        )}

        {step === 'subfolders' && sourceFolder && (
          <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface }}>
            <div style={{ padding: '12px 14px', borderBottom: `1px solid ${t.border}` }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: t.text }}>Select subfolders and/or files to process</div>
              <div style={{ fontSize: 11, color: t.textSubtle, marginTop: 2 }}>
                Only what's checked below is scanned — everything else in {sourceFolder.name} is left alone. Checked subfolders are scanned fully (including everything nested).
              </div>
            </div>
            {subfolderLoading ? (
              <div style={{ padding: 30, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>Loading…</div>
            ) : (
              <>
                {subfolderOptions.length === 0 && subfolderFiles.length === 0 && (
                  <div style={{ padding: '10px 14px', fontSize: 12, color: t.textSubtle }}>This folder is empty.</div>
                )}
                {subfolderOptions.map((f) => (
                  <label key={f.id} style={{ ...rowStyle, cursor: 'pointer' }}>
                    <input type="checkbox" checked={checkedSubfolders.has(f.name)} onChange={() => toggleSubfolder(f.name)} />
                    <Icon iconName="FabricFolder" style={{ color: t.textSubtle }} />
                    <span style={{ color: t.text, fontSize: 13 }}>{f.name}</span>
                  </label>
                ))}
                {subfolderFiles.map((f) => (
                  <label key={f.id} style={{ ...rowStyle, cursor: 'pointer' }}>
                    <input type="checkbox" checked={checkedFiles.has(f.name)} onChange={() => toggleFile(f.name)} />
                    <Icon iconName={fileIconName(f.name.split('.').pop()).iconName} style={{ color: fileIconName(f.name.split('.').pop()).color }} />
                    <span style={{ flex: 1, color: t.text, fontSize: 13 }}>{f.name}</span>
                    <span style={{ fontSize: 11, color: t.textSubtle }}>{formatSize(f.size)}</span>
                  </label>
                ))}
              </>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 14px', borderTop: `1px solid ${t.border}` }}>
              <button style={secondaryBtn(isNight)} onClick={() => setStep('source')}>Back</button>
              <button
                style={primaryBtn(subfolderOptions.length > 0 && checkedSubfolders.size === 0 && checkedFiles.size === 0)}
                disabled={subfolderOptions.length > 0 && checkedSubfolders.size === 0 && checkedFiles.size === 0}
                onClick={goToVesselStep}
              >
                {checkedSubfolders.size === 0 && checkedFiles.size === 0 ? 'Next: pick destination vessel (scan this folder directly)' : `Next: pick destination vessel (${checkedSubfolders.size + checkedFiles.size} selected)`}
              </button>
            </div>
          </div>
        )}

        {step === 'vessel' && (
          <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface }}>
            <div style={{ padding: '12px 14px', borderBottom: `1px solid ${t.border}` }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: t.text }}>Select the destination vessel</div>
              <div style={{ fontSize: 11, color: t.textSubtle, marginTop: 2 }}>Documents will be filed into this vessel's existing folders.</div>
            </div>
            {!vesselLoading && destinationRoot && (
              <div style={{ ...rowStyle, background: 'var(--clay-accent-soft)' }}>
                <Icon iconName="Sparkle" style={{ color: 'var(--clay-accent)' }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: t.text }}>Auto-detect vessel</div>
                  <div style={{ fontSize: 11, color: t.textSubtle }}>Each file is routed to the vessel it actually belongs to.</div>
                </div>
                <button style={primaryBtn()} onClick={() => startScan({ id: '', name: destinationRoot, path: destinationRoot })}>Scan into {destinationRoot}</button>
              </div>
            )}
            {vesselLoading ? (
              <div style={{ padding: 30, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>Loading vessels…</div>
            ) : vessels.length === 0 ? (
              <div style={{ padding: 30, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>No vessel folders found.</div>
            ) : (
              vessels.map((v) => (
                <div key={v.id} style={rowStyle}>
                  <Icon iconName="FabricFolder" style={{ color: t.textSubtle }} />
                  <span style={{ flex: 1, color: t.text, fontSize: 13 }}>{v.name}</span>
                  <button style={primaryBtn()} onClick={() => startScan(v)}>Scan into this vessel</button>
                </div>
              ))
            )}
            <div style={{ padding: '12px 14px', borderTop: `1px solid ${t.border}` }}>
              <button style={secondaryBtn(isNight)} onClick={() => setStep('subfolders')}>Back</button>
            </div>
          </div>
        )}

        {step === 'scanning' && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10, padding: 60, border: `1px solid ${t.border}`, borderRadius: 14, background: t.surface }}>
            <Icon iconName="Sync" style={{ fontSize: 24, color: 'var(--clay-accent)' }} />
            <div style={{ fontSize: 13, color: t.textMuted }}>
              {job && job.total_found > 0 ? `Classifying documents (${job.processed} of ${job.total_found})…` : 'Scanning for documents…'}
            </div>
          </div>
        )}

        {step === 'preview' && job && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface, padding: 14 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: t.text }}>
                {items.length} document(s) found in {[...job.subfolders, ...job.files].join(', ') || job.source_folder} — destination:{' '}
                <span style={{ color: 'var(--clay-accent)' }}>{job.auto_detect_vessel ? `auto-detect (${job.vessel_name})` : job.vessel_name}</span>
              </div>
              <div style={{ fontSize: 11, color: t.textSubtle, marginTop: 4 }}>Review each row's category below. Nothing moves until "Confirm Move."</div>
            </div>

            <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, overflow: 'hidden' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <thead style={{ background: t.surfaceAlt }}>
                  <tr>
                    <th style={{ textAlign: 'left', padding: '8px 12px', fontSize: 11, color: t.textMuted, textTransform: 'uppercase' }}>File</th>
                    {job.auto_detect_vessel && <th style={{ textAlign: 'left', padding: '8px 12px', fontSize: 11, color: t.textMuted, textTransform: 'uppercase' }}>Vessel</th>}
                    <th style={{ textAlign: 'left', padding: '8px 12px', fontSize: 11, color: t.textMuted, textTransform: 'uppercase' }}>AI Category</th>
                    <th style={{ textAlign: 'left', padding: '8px 12px', fontSize: 11, color: t.textMuted, textTransform: 'uppercase' }}>Destination</th>
                    <th style={{ textAlign: 'left', padding: '8px 12px', fontSize: 11, color: t.textMuted, textTransform: 'uppercase' }}>Confidence</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {items.map((item) => {
                    const meta = fileIconName(item.filename.split('.').pop());
                    const conf = confidenceLabel(item);
                    const vesselLabel = job.auto_detect_vessel ? item.detected_vessel_name ?? job.vessel_name : job.vessel_name;
                    const destination = item.suggested_path ? `${vesselLabel}/${item.suggested_path}` : 'Category tag required';
                    return (
                      <tr key={item.id} style={{ borderTop: `1px solid ${t.border}` }}>
                        <td style={{ padding: '8px 12px', maxWidth: 220 }}>
                          <span style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer' }} onClick={() => setPreviewItem(item)}>
                            <Icon iconName={meta.iconName} style={{ color: meta.color, fontSize: 14 }} />
                            <span style={{ color: t.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.filename}</span>
                          </span>
                        </td>
                        {job.auto_detect_vessel && (
                          <td style={{ padding: '8px 12px', fontSize: 12, color: t.text }}>
                            {item.detected_vessel_name || <span style={{ color: t.textMuted }}>unknown</span>}
                            {item.vessel_exists === false && <span style={{ ...pill('primary'), marginLeft: 6 }}>new</span>}
                          </td>
                        )}
                        <td style={{ padding: '8px 12px' }}>
                          {item.category ? (
                            <span style={pill('primary')}>{item.category}{item.subcategory ? ` / ${item.subcategory}` : ''}</span>
                          ) : item.status === 'failed' ? (
                            <span style={pill('error')} title={item.error ?? 'Classification failed'}>Failed</span>
                          ) : (
                            <span style={pill('warning')}>Needs review</span>
                          )}
                        </td>
                        <td style={{ padding: '8px 12px' }}>
                          <select
                            value={item.suggested_path ?? ''}
                            disabled={busyItemId === item.id}
                            onChange={(e) => handleOverride(item, e.target.value)}
                            style={{ maxWidth: 200, borderRadius: 8, border: `1px solid ${t.border}`, background: t.surface, color: t.text, fontSize: 12, padding: '4px 6px' }}
                          >
                            <option value="" disabled>{item.suggested_path ? destination : 'Category tag required'}</option>
                            {hierarchyPaths.map((p) => <option key={p} value={p}>{vesselLabel}/{p}</option>)}
                          </select>
                        </td>
                        <td style={{ padding: '8px 12px', color: conf.color, fontWeight: 600 }}>{conf.text}</td>
                        <td style={{ padding: '8px 12px', textAlign: 'right' }}>
                          <button
                            disabled={busyItemId === item.id}
                            onClick={() => handleReclassify(item)}
                            title="Re-run AI classification"
                            style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: t.textMuted }}
                          >
                            <Icon iconName="Sync" style={{ fontSize: 13 }} />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <button style={secondaryBtn(isNight)} onClick={() => setStep('vessel')}>Back</button>
              <button style={{ ...primaryBtn(confirming || items.length === 0), background: confirming || items.length === 0 ? undefined : `linear-gradient(150deg, #34d399, #16a34a)` }} disabled={confirming || items.length === 0} onClick={handleConfirm}>
                <Icon iconName="CheckMark" style={{ fontSize: 13 }} />Confirm Move
              </button>
            </div>
          </div>
        )}

        {step === 'summary' && summary && job && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface, padding: 20 }}>
              <div style={{ fontSize: 16, fontWeight: 700, color: t.text }}>Move complete</div>
              <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>Destination vessel: {job.vessel_name}</div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 10, marginTop: 14 }}>
                <SummaryStat label="Total scanned" value={summary.total} isNight={isNight} />
                <SummaryStat label="Categorized" value={summary.categorized} isNight={isNight} />
                <SummaryStat label="Moved" value={summary.moved} color={successColor} isNight={isNight} />
                <SummaryStat label="Failed" value={summary.failed} color={errColor} isNight={isNight} />
              </div>
            </div>
            {summary.failed_items.length > 0 && (
              <div style={{ borderRadius: 14, border: `1px solid ${errColor}55`, background: `${errColor}10`, padding: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: errColor }}>Failed files</div>
                {summary.failed_items.map((f, i) => (
                  <div key={i} style={{ fontSize: 12, color: errColor, marginTop: 4 }}><b>{f.filename}</b> — {f.error}</div>
                ))}
              </div>
            )}
            <button style={primaryBtn()} onClick={startOver}>Start another migration</button>
          </div>
        )}
      </div>

      {previewItem && job && <FilePreviewDrawer api={api} item={previewItem} jobId={job.id} isNight={isNight} actingEmail={actingEmail} onClose={() => setPreviewItem(null)} />}
      {recentOpen && <RecentJobsPicker jobs={recentJobs} loading={recentLoading} isNight={isNight} onClose={() => setRecentOpen(false)} onPick={viewJob} />}
    </div>
  );
}

function SummaryStat({ label, value, color, isNight }: { label: string; value: number; color?: string; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  return (
    <div style={{ borderRadius: 10, border: `1px solid ${t.border}`, background: t.surfaceAlt, padding: '10px 8px', textAlign: 'center' }}>
      <div style={{ fontSize: 18, fontWeight: 700, color: color || t.text }}>{value}</div>
      <div style={{ fontSize: 11, color: t.textMuted, marginTop: 2 }}>{label}</div>
    </div>
  );
}

function FileRow({ file, isNight }: { file: SourceFile; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  const meta = fileIconName(file.name.split('.').pop());
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', opacity: 0.8, borderBottom: `1px solid ${t.border}` }}>
      <Icon iconName={meta.iconName} style={{ color: meta.color }} />
      <span style={{ flex: 1, color: t.text, fontSize: 13 }}>{file.name}</span>
      <span style={{ fontSize: 11, color: t.textSubtle }}>{formatSize(file.size)}</span>
    </div>
  );
}

function RecentJobsPicker({ jobs, loading, isNight, onClose, onPick }: {
  jobs: MigrationScanJob[]; loading: boolean; isNight: boolean; onClose: () => void; onPick: (j: MigrationScanJob) => void;
}): React.ReactElement {
  const t = tokens(isNight);
  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 1000, background: 'rgba(0,0,0,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center' }} onClick={onClose}>
      <div style={{ width: '100%', maxWidth: 560, maxHeight: '80vh', display: 'flex', flexDirection: 'column', borderRadius: 14, background: t.surface, boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 16px', borderBottom: `1px solid ${t.border}` }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: t.text }}>Recent scans</span>
          <button onClick={onClose} style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: t.textMuted }}><Icon iconName="Cancel" /></button>
        </div>
        <div style={{ flex: 1, overflowY: 'auto', padding: 8 }}>
          {loading ? <div style={{ padding: 20, textAlign: 'center', color: t.textMuted }}>Loading…</div> : jobs.length === 0 ? (
            <div style={{ padding: 16, fontSize: 13, color: t.textMuted }}>No scans yet.</div>
          ) : jobs.map((j) => (
            <div key={j.id} onClick={() => onPick(j)} style={{ display: 'flex', gap: 10, padding: '10px 12px', borderRadius: 10, cursor: 'pointer' }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: t.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {[...j.subfolders, ...j.files].length > 0 ? `${j.source_folder}/{${[...j.subfolders, ...j.files].join(', ')}}` : j.source_folder}
                </div>
                <div style={{ fontSize: 11, color: t.textSubtle }}>→ {j.vessel_name} · {j.total_found} document(s){j.confirmed_at ? ' · confirmed' : ''}</div>
              </div>
              <span style={pill(j.status === 'running' ? 'warning' : j.status === 'failed' ? 'error' : 'success')}>{j.status}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function FilePreviewDrawer({ api, item, jobId, isNight, actingEmail, onClose }: {
  api: MigrationApi; item: MigrationItem; jobId: string; isNight: boolean; actingEmail: string; onClose: () => void;
}): React.ReactElement {
  const t = tokens(isNight);
  const [blobUrl, setBlobUrl] = React.useState<string | null>(null);
  const [loadError, setLoadError] = React.useState(false);
  const meta = fileIconName(item.filename.split('.').pop());

  React.useEffect(() => {
    let active = true;
    let objUrl: string | null = null;
    setBlobUrl(null);
    setLoadError(false);
    fetch(api.jobItemPreviewUrl(jobId, item.id), { headers: api.authHeaders() })
      .then((r) => (r.ok ? r.blob() : Promise.reject(new Error('fail'))))
      .then((blob) => {
        if (!active) return;
        objUrl = URL.createObjectURL(blob);
        setBlobUrl(objUrl);
      })
      .catch(() => active && setLoadError(true));
    return () => {
      active = false;
      if (objUrl) URL.revokeObjectURL(objUrl);
    };
  }, [api, jobId, item.id, actingEmail]);

  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 1000, background: 'rgba(0,0,0,0.4)', display: 'flex', justifyContent: 'flex-end' }} onClick={onClose}>
      <div style={{ width: '100%', maxWidth: 560, height: '100%', display: 'flex', flexDirection: 'column', background: t.surface }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '12px 16px', borderBottom: `1px solid ${t.border}` }}>
          <Icon iconName={meta.iconName} style={{ color: meta.color }} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: t.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.filename}</div>
            <div style={{ fontSize: 11, color: t.textSubtle }}>{[meta.label, formatSize(item.size)].filter(Boolean).join(' · ')}</div>
          </div>
          <button onClick={onClose} style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: t.textMuted }}><Icon iconName="Cancel" /></button>
        </div>
        {item.reason && (
          <div style={{ padding: '10px 16px', borderBottom: `1px solid ${t.border}`, fontSize: 12 }}>
            <span style={{ fontWeight: 600, color: t.text }}>AI reasoning: </span><span style={{ color: t.textMuted }}>{item.reason}</span>
            {item.keywords.length > 0 && (
              <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {item.keywords.map((k) => <span key={k} style={{ ...pill('primary') }}>{k}</span>)}
              </div>
            )}
          </div>
        )}
        <div style={{ flex: 1, overflow: 'auto', background: t.surfaceAlt }}>
          {loadError ? (
            <div style={{ display: 'flex', height: '100%', alignItems: 'center', justifyContent: 'center', fontSize: 13, color: t.textMuted }}>Couldn't load a preview for this document.</div>
          ) : !blobUrl ? (
            <div style={{ display: 'flex', height: '100%', alignItems: 'center', justifyContent: 'center' }}><Icon iconName="Sync" style={{ fontSize: 20, color: t.textSubtle }} /></div>
          ) : meta.previewable ? (
            item.filename.toLowerCase().endsWith('.pdf') ? (
              <iframe title={item.filename} src={blobUrl} style={{ width: '100%', height: '100%', border: 0 }} />
            ) : (
              <div style={{ display: 'flex', height: '100%', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
                <img src={blobUrl} alt={item.filename} style={{ maxWidth: '100%', maxHeight: '100%', borderRadius: 6 }} />
              </div>
            )
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', height: '100%', alignItems: 'center', justifyContent: 'center', gap: 8 }}>
              <Icon iconName={meta.iconName} style={{ fontSize: 42, color: meta.color }} />
              <div style={{ fontSize: 12, color: t.textMuted }}>Preview isn't available for {meta.label} files.</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
