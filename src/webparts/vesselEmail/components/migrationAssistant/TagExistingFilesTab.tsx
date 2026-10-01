// Ported from the standalone project's TagExistingFiles.tsx.
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { MigrationApi, errDetail } from './api';
import { pill, primaryBtn, secondaryBtn, tokens, error as errColor } from './styles';
import type { ExistingFileTagApply, ExistingFileTagRow, ExistingFileTagScan, TaggedScanJob } from './types';

function TagValue({ value }: { value: string | null | undefined }): React.ReactElement {
  return <span style={{ color: value ? undefined : errColor }}>{value || 'blank'}</span>;
}

export function TagExistingFilesTab({ api, isNight }: { api: MigrationApi; isNight: boolean }): React.ReactElement {
  const t = tokens(isNight);
  const [rootPath, setRootPath] = React.useState('Technical and Crewing');
  const [scan, setScan] = React.useState<ExistingFileTagScan | null>(null);
  const [selected, setSelected] = React.useState<Set<string>>(new Set());
  const [result, setResult] = React.useState<ExistingFileTagApply | null>(null);
  const [busy, setBusy] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | null>(null);
  const [recentOpen, setRecentOpen] = React.useState(false);
  const [recentJobs, setRecentJobs] = React.useState<TaggedScanJob[]>([]);
  const [recentLoading, setRecentLoading] = React.useState(false);

  const runScan = async (): Promise<void> => {
    setBusy(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const data = await api.scanExistingFiles(rootPath);
      setScan(data);
      setSelected(new Set(data.files.filter((f) => f.ready).map((f) => f.id)));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not scan destination files.'));
    } finally {
      setBusy(false);
    }
  };

  const openRecentJobs = async (): Promise<void> => {
    setRecentOpen(true);
    setRecentLoading(true);
    try {
      setRecentJobs(await api.listTaggedScanJobs());
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load recent tagged scans.'));
    } finally {
      setRecentLoading(false);
    }
  };

  const viewJob = async (job: TaggedScanJob): Promise<void> => {
    setRecentOpen(false);
    try {
      const data = await api.getTaggedScanJob(job.id);
      setRootPath(data.root_path);
      setScan(data);
      setSelected(new Set(data.files.filter((f) => f.ready).map((f) => f.id)));
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not load that tagged scan.'));
    }
  };

  const toggle = (id: string): void =>
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });

  const apply = async (): Promise<void> => {
    setBusy(true);
    setErrorMsg(null);
    try {
      setResult(await api.applyExistingFileTags(rootPath, Array.from(selected)));
      await runScan();
    } catch (e) {
      setErrorMsg(errDetail(e, 'Could not apply tags.'));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ padding: '14px 4px' }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: t.text, display: 'flex', alignItems: 'center', gap: 6 }}>
          <Icon iconName="Tag" style={{ color: 'var(--clay-accent)' }} />Tag Existing Files
        </div>
        <div style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>
          Derives Group, Category and Vessel only from a file's current folder path — no move, no content extraction. Only writes Managed Metadata columns.
        </div>
        <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
          <input value={rootPath} onChange={(e) => setRootPath(e.target.value)} placeholder="Destination root folder path"
            style={{ minWidth: 260, borderRadius: 8, border: `1px solid ${t.border}`, background: t.surface, color: t.text, padding: '7px 10px', fontSize: 13 }} />
          <button style={primaryBtn(busy)} disabled={busy} onClick={runScan}><Icon iconName="Sync" style={{ fontSize: 12 }} />Scan destination</button>
          <button style={secondaryBtn(isNight)} onClick={openRecentJobs}><Icon iconName="History" style={{ fontSize: 12 }} />Recent tagged scans</button>
          <button style={primaryBtn(busy || selected.size === 0 || !scan)} disabled={busy || selected.size === 0 || !scan} onClick={apply}>
            <Icon iconName="CheckMark" style={{ fontSize: 12 }} />Apply Tags ({selected.size})
          </button>
        </div>
      </div>

      {errorMsg && <div style={{ borderRadius: 10, border: `1px solid ${errColor}55`, background: `${errColor}15`, color: errColor, padding: '8px 14px', fontSize: 13, marginBottom: 10 }}>{errorMsg}</div>}

      {scan && (
        <div style={{ flex: 1, overflowY: 'auto' }}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 8, marginBottom: 12 }}>
            <Stat label="Files found" value={scan.summary.total_files} isNight={isNight} />
            <Stat label="Fully tagged" value={scan.summary.fully_tagged} isNight={isNight} color="#16a34a" />
            <Stat label="Category skipped" value={scan.summary.category_skipped} isNight={isNight} color="#d97706" />
            <Stat label="Missing taxonomy" value={scan.summary.missing_taxonomy} isNight={isNight} />
            <Stat label="Failed" value={scan.summary.missing_terms} isNight={isNight} color={errColor} />
          </div>
          <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, overflow: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead style={{ background: t.surfaceAlt }}>
                <tr>
                  <th style={{ padding: 8 }} />
                  <th style={{ textAlign: 'left', padding: 8, color: t.textMuted }}>File / folder</th>
                  <th style={{ textAlign: 'left', padding: 8, color: t.textMuted }}>Current</th>
                  <th style={{ textAlign: 'left', padding: 8, color: t.textMuted }}>Folder-derived</th>
                  <th style={{ textAlign: 'left', padding: 8, color: t.textMuted }}>Terms</th>
                </tr>
              </thead>
              <tbody>
                {scan.files.map((file: ExistingFileTagRow) => (
                  <tr key={file.id} style={{ borderTop: `1px solid ${t.border}` }}>
                    <td style={{ padding: 8 }}><input type="checkbox" checked={selected.has(file.id)} disabled={!file.ready} onChange={() => toggle(file.id)} /></td>
                    <td style={{ padding: 8 }}>
                      <div style={{ fontWeight: 600, color: t.text }}>{file.filename}</div>
                      <div style={{ fontSize: 11, color: t.textSubtle }}>{file.folder_path}</div>
                    </td>
                    <td style={{ padding: 8, lineHeight: 1.6 }}>G: <TagValue value={file.current.group} /><br />C: <TagValue value={file.current.category} /><br />V: <TagValue value={file.current.vessel} /></td>
                    <td style={{ padding: 8, lineHeight: 1.6 }}>
                      {file.derived ? <>G: <TagValue value={file.derived.group} /><br />C: <TagValue value={file.derived.category} /><br />V: <TagValue value={file.derived.vessel} /></> : <span style={{ color: errColor }}>Unable to determine tags</span>}
                    </td>
                    <td style={{ padding: 8, lineHeight: 1.6 }}>
                      {file.derived ? Object.entries(file.terms).map(([key, term]) => {
                        const skippedOptional = !term.resolved && key === 'category' && file.ready;
                        const color = term.resolved ? '#16a34a' : skippedOptional ? '#d97706' : errColor;
                        return <div key={key} title={term.reason} style={{ color }}>{key}: {term.resolved ? term.guid : skippedOptional ? `skipped — ${term.reason}` : term.reason}</div>;
                      }) : <span style={{ color: errColor }}>Not ready</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {result && (
        <div style={{ borderTop: `1px solid ${t.border}`, padding: '10px 4px', fontSize: 13, color: t.text }}>
          Fully tagged: <b style={{ color: '#16a34a' }}>{result.fully_tagged}</b> · Tagged (category skipped): <b style={{ color: '#d97706' }}>{result.category_skipped}</b> · Failed: <b style={{ color: errColor }}>{result.failed}</b>
        </div>
      )}

      {recentOpen && (
        <div style={{ position: 'fixed', inset: 0, zIndex: 1000, background: 'rgba(0,0,0,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center' }} onClick={() => setRecentOpen(false)}>
          <div style={{ width: '100%', maxWidth: 560, maxHeight: '80vh', overflowY: 'auto', borderRadius: 14, background: t.surface, padding: 8 }} onClick={(e) => e.stopPropagation()}>
            <div style={{ fontSize: 13, fontWeight: 700, color: t.text, padding: '8px 8px 12px' }}>Recent tagged scans</div>
            {recentLoading ? <div style={{ padding: 16 }}>Loading…</div> : recentJobs.length === 0 ? <div style={{ padding: 16, color: t.textMuted, fontSize: 13 }}>No tagged scans yet.</div> : recentJobs.map((job) => (
              <div key={job.id} onClick={() => viewJob(job)} style={{ padding: '10px 8px', cursor: 'pointer', display: 'flex', justifyContent: 'space-between', gap: 8 }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: t.text }}>{job.root_path || 'Destination root'}</div>
                  <div style={{ fontSize: 11, color: t.textSubtle }}>{job.summary.total_files} files · {job.summary.fully_tagged} fully tagged</div>
                </div>
                <span style={pill('success')}>{job.status}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function Stat({ label, value, isNight, color }: { label: string; value: number; isNight: boolean; color?: string }): React.ReactElement {
  const t = tokens(isNight);
  return (
    <div style={{ borderRadius: 10, border: `1px solid ${t.border}`, background: t.surface, padding: '10px 8px' }}>
      <div style={{ fontSize: 11, color: t.textMuted }}>{label}</div>
      <div style={{ fontSize: 18, fontWeight: 700, color: color || t.text }}>{value}</div>
    </div>
  );
}
