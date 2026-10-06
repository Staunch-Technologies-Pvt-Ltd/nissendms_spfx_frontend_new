import * as React from 'react';
import type VesselEmail from './VesselEmail';
import type { FlatRow } from './types/rows';
import type { VesselSuggestionUploadEntry } from './types/ui';
import { isMobileWidth } from './responsive';
import { Icon } from '@fluentui/react/lib/Icon';

// ── Types ────────────────────────────────────────────────────────────────────

export interface BulkUploadFile {
  file: File;
  /** Relative path inside an uploaded folder, e.g. "Reports/Jan/invoice.pdf" */
  relativePath?: string;
}

export interface BulkUploadProgress {
  id: string;
  name: string;
  relativePath?: string;
  status: 'queued' | 'uploading' | 'done' | 'failed';
  message?: string;
  size: string;
  /** Direct SharePoint link to the uploaded file, once known — lets you verify the real destination. */
  webUrl?: string | null;
}

export interface BulkUploadModalProps {
  host: VesselEmail;
  files: BulkUploadFile[];
  /** SPO drive item ID of the target folder */
  folderId: string;
  subFolderPath: string;
  vesselName: string;
  currentFolderNode: { id: string; name: string } | null;
  targetSiteId?: string;
  targetDriveId?: string;
  onClose: () => void;
}

type TagKey = 'domain' | 'vessel' | 'group' | 'category';
type TagValues = Record<TagKey, string>;
type TagOptions = Record<TagKey, string[]>;

const TAG_FIELDS: Array<{ key: TagKey; label: string }> = [
  { key: 'domain', label: 'Domain' },
  { key: 'vessel', label: 'Vessel' },
  { key: 'group', label: 'Group' },
  { key: 'category', label: 'Category' },
];
const EMPTY_TAGS: TagValues = { domain: '', vessel: '', group: '', category: '' };

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Loose match key: case-insensitive, '&' == 'and', punctuation ignored. */
function normTag(v: string): string {
  return (v || '').toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, '');
}

/** First option that equals one of the path segments (or the fallback value). */
function pickTag(options: string[], segments: string[], fallback?: string): string {
  for (const seg of segments) {
    const hit = options.find(o => normTag(o) === normTag(seg));
    if (hit) return hit;
  }
  if (fallback) {
    const hit = options.find(o => normTag(o) === normTag(fallback));
    if (hit) return hit;
  }
  return '';
}

function fmtSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

// ── Component ────────────────────────────────────────────────────────────────

export function BulkUploadModal({
  host, files, folderId, subFolderPath, vesselName, currentFolderNode, targetSiteId, targetDriveId, onClose,
}: BulkUploadModalProps): React.ReactElement {
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  const [progress, setProgress] = React.useState<BulkUploadProgress[]>(() =>
    files.map((f, i) => ({
      id: `bulk_${i}_${f.file.name}`,
      name: f.file.name,
      relativePath: f.relativePath,
      status: 'queued',
      size: fmtSize(f.file.size),
    }))
  );

  const [started, setStarted] = React.useState(false);
  const [done, setDone] = React.useState(false);
  const [elapsedSeconds, setElapsedSeconds] = React.useState(0);
  const [autoCloseSeconds, setAutoCloseSeconds] = React.useState<number | null>(null);
  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [lastSuccessWebUrl, setLastSuccessWebUrl] = React.useState<string | null>(null);
  const hasRunRef = React.useRef(false);

  // Tag review step. Pre-fills Domain / Vessel / Group / Category from the
  // destination folder path. Nothing is sent unless the user changes a value;
  // if the vocabulary can't be loaded the dialog starts uploading immediately,
  // exactly as it did before this step existed.
  const [reviewing, setReviewing] = React.useState(true);
  const [tagOptions, setTagOptions] = React.useState<TagOptions | null>(null);
  const [tagError, setTagError] = React.useState('');
  const [tags, setTags] = React.useState<TagValues>(EMPTY_TAGS);
  const autoTagsRef = React.useRef<TagValues>(EMPTY_TAGS);

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const base = host._base();
        const headers = host._headers();
        const [activeRes, vesselRes] = await Promise.all([
          fetch(`${base}/api/tag-config/active`, { headers }),
          fetch(`${base}/api/tag-config/vessel-names`, { headers }).catch(() => null),
        ]);
        if (!activeRes.ok) throw new Error(`HTTP ${activeRes.status}`);
        const active = await activeRes.json() as { tree?: Array<{ level: string; name: string; display_name?: string }>; domain_aliases?: Record<string, string[]> };
        const rows = active.tree || [];
        const namesFor = (lvl: string): string[] =>
          Array.from(new Set(rows.filter(r => r.level === lvl).map(r => r.display_name || r.name))).sort((a, b) => a.localeCompare(b));
        let vessels: string[] = [];
        if (vesselRes && vesselRes.ok) {
          const vj = await vesselRes.json() as { vessel_names?: string[] };
          vessels = (vj.vessel_names || []).slice().sort((a, b) => a.localeCompare(b));
        }
        if (cancelled) return;
        const segments = (subFolderPath || '').split(/[>/]/).map(x => x.trim()).filter(Boolean);
        // Never leave Vessel empty-handed: if the Term Store list is empty or
        // doesn't contain the vessel we're uploading into, offer that vessel.
        if (vesselName && !vessels.some(v => normTag(v) === normTag(vesselName))) vessels = [vesselName, ...vessels];
        const options: TagOptions = { domain: namesFor('domain'), vessel: vessels, group: namesFor('group'), category: namesFor('category') };
        const auto: TagValues = {
          domain: pickTag(options.domain, segments)
            || (options.domain.find(d => segments.some(seg =>
              (active.domain_aliases?.[d] || []).some(a => normTag(a) === normTag(seg)))) || ''),
          vessel: pickTag(options.vessel, segments, vesselName),
          group: pickTag(options.group, segments),
          category: pickTag(options.category, segments),
        };
        autoTagsRef.current = auto;
        console.info('[VesselDMS] BulkUploadModal tag options loaded', { auto, counts: { domain: options.domain.length, vessel: options.vessel.length, group: options.group.length, category: options.category.length } });
        setTagOptions(options);
        setTags(auto);
      } catch (e) {
        console.warn('[VesselDMS] Tag options unavailable.', e);
        if (!cancelled) {
          setTagError((e as Error)?.message || 'Could not load the tag lists.');
          setTagOptions({ domain: [], vessel: vesselName ? [vesselName] : [], group: [], category: [] });
        }
      }
    })();
    return () => { cancelled = true; };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  React.useEffect(() => {
    console.info('[VesselDMS] BulkUploadModal mounted', { files: files.length, subFolderPath, vesselName });
    return () => { console.info('[VesselDMS] BulkUploadModal unmounted'); };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const tagsEdited = TAG_FIELDS.some(f => tags[f.key] !== autoTagsRef.current[f.key]);

  // Elapsed timer while uploading
  React.useEffect(() => {
    if (!started || done) return undefined;
    const iv = window.setInterval(() => setElapsedSeconds(s => s + 1), 1000);
    return () => window.clearInterval(iv);
  }, [started, done]);

  // Auto-close countdown after completion
  React.useEffect(() => {
    if (autoCloseSeconds === null) return undefined;
    if (autoCloseSeconds <= 0) { onClose(); return undefined; }
    const t = window.setTimeout(() => setAutoCloseSeconds(v => v === null ? null : v - 1), 1000);
    return () => window.clearTimeout(t);
  }, [autoCloseSeconds]);

  const updateProgress = (id: string, update: Partial<BulkUploadProgress>): void => {
    setProgress(prev => prev.map(p => p.id === id ? { ...p, ...update } : p));
  };

  // Run upload once on mount
  React.useEffect(() => {
    if (reviewing || hasRunRef.current) return;
    hasRunRef.current = true;
    setStarted(true);
    const overrideTags = tagsEdited ? { ...tags } : null;

    void (async () => {
      const newUploads: Array<{ name: string; size: string; date: string; id?: string; uploadedAt: number }> = [];
      const successFilesList: File[] = []; // tracks files that uploaded OK — used for vessel suggestions
      const successUploadEntries: VesselSuggestionUploadEntry[] = [];
      let lastWebUrl: string | null = null;

      for (let i = 0; i < files.length; i++) {
        const { file, relativePath } = files[i];
        const progressId = `bulk_${i}_${file.name}`;
        setCurrentIndex(i);
        updateProgress(progressId, { status: 'uploading' });

        let targetSubFolderPath = subFolderPath;
        if (relativePath && relativePath.includes('/')) {
          const relDir = relativePath.substring(0, relativePath.lastIndexOf('/'));
          const extraSegments = relDir.split('/').filter(Boolean).join(' > ');
          targetSubFolderPath = subFolderPath
            ? `${subFolderPath} > ${extraSegments}`
            : extraSegments;
        }

        const optimisticKeys = Array.from(new Set([
          targetSubFolderPath,
          (targetSubFolderPath || '').toLowerCase(),
          subFolderPath,
          (subFolderPath || '').toLowerCase(),
          currentFolderNode?.id,
          currentFolderNode?.name,
          folderId,
        ].filter(Boolean) as string[]));
        const optimisticId = host._addUploadingFilePlaceholder(optimisticKeys, file);

        try {
          // Use host._uploadFileToFolder which reliably resolves SharePoint folder paths, Graph IDs, and REST fallbacks
          const result = await host._uploadFileToFolder(
            folderId,
            targetSubFolderPath,
            vesselName,
            file,
            false,
            targetSiteId,
            targetDriveId,
            subFolderPath
          );

          host._replaceUploadingFilePlaceholder(optimisticKeys, optimisticId, file, {
            id: result.fileId,
            pending: result.statusPending,
            date: result.statusPending ? 'Queued for approval' : 'Just now',
            uploadedAt: Date.now(),
          });

          let tagNote = '';
          if (overrideTags && result.isGraphUpload && result.fileId) {
            try {
              const effSite = targetSiteId || host.props.siteId;
              const effDrive = targetDriveId || host.props.driveId;
              const res = await fetch(
                `${host._base()}/api/sites/${encodeURIComponent(effSite || '')}/drives/${encodeURIComponent(effDrive || '')}/items/${encodeURIComponent(result.fileId)}/tags`,
                {
                  method: 'PATCH',
                  headers: host._headers(),
                  body: JSON.stringify({
                    department: overrideTags.domain, vessel: overrideTags.vessel,
                    group: overrideTags.group, category: overrideTags.category,
                  }),
                },
              );
              const body = await res.json().catch(() => ({})) as { detail?: string; ok?: boolean; metadata_patch?: { error?: string } };
              if (!res.ok) tagNote = ` (uploaded, but tags not saved: ${body.detail || `HTTP ${res.status}`})`;
              else if (body.ok === false) tagNote = ` (uploaded, but tags not saved: ${body.metadata_patch?.error || 'SharePoint rejected them'})`;
            } catch (tagErr: any) {
              tagNote = ` (uploaded, but tags not saved: ${tagErr?.message || 'network error'})`;
            }
          }

          const fileId = result.fileId || `file_${Date.now()}_${i}`;
          const newUpload = {
            name: file.name,
            size: fmtSize(file.size),
            date: 'Just now',
            id: fileId,
            uploadedAt: Date.now(),
          };
          newUploads.push(newUpload);
          successFilesList.push(file); // track for vessel suggestions
          successUploadEntries.push({
            filename: file.name,
            drive_item_id: result.fileId || null,
            source_subfolder_path: targetSubFolderPath || null,
            source_vessel_name: vesselName || null,
            uploaded_at: Date.now(),
          });

          // Update state immediately for this file
          host.setState(prev => {
            const uploadedFilesByFolder = { ...prev.uploadedFilesByFolder };
            const baseRows = (prev.rows || []) as FlatRow[];

            // Find all rows matching this vessel + folder
            const vLower = (vesselName || '').toLowerCase();
            const folderNorm = (currentFolderNode?.name || '').toLowerCase();
            const subPathNorm = (subFolderPath || '').toLowerCase();
            const targetNorm = (targetSubFolderPath || '').toLowerCase();

            const matchingRows = baseRows.filter(r =>
              r.vesselName?.toLowerCase() === vLower &&
              (
                (folderNorm && (r.subCategory?.toLowerCase() === folderNorm || r.category?.toLowerCase() === folderNorm)) ||
                (r.subFolderPath && (r.subFolderPath.toLowerCase() === subPathNorm || r.subFolderPath.toLowerCase() === targetNorm)) ||
                (result.folderId && r.uploadFolderId === result.folderId)
              )
            );

            const isDeepUpload = !!targetSubFolderPath && !!subFolderPath &&
              targetSubFolderPath.trim().toLowerCase() !== subFolderPath.trim().toLowerCase();

            // Store file under all relevant keys so folder & list views see it immediately in its destination folder
            const keysToSet = new Set<string>([
              targetSubFolderPath,
              targetSubFolderPath.toLowerCase(),
              result.folderId,
              ...(!isDeepUpload ? [
                currentFolderNode?.id,
                currentFolderNode?.name,
                subFolderPath,
                subFolderPath.toLowerCase(),
              ] : []),
            ].filter(Boolean) as string[]);

            for (const mr of matchingRows) {
              if (isDeepUpload) {
                const mrPath = (mr.subFolderPath || '').trim().toLowerCase();
                if (mrPath === targetNorm) {
                  if (mr.groupKey) keysToSet.add(mr.groupKey);
                  if (mr.uploadFolderId) keysToSet.add(mr.uploadFolderId);
                }
              } else {
                if (mr.groupKey) keysToSet.add(mr.groupKey);
                if (mr.subFolderPath) {
                  keysToSet.add(mr.subFolderPath);
                  keysToSet.add(mr.subFolderPath.toLowerCase());
                }
                if (mr.subCategory) keysToSet.add(mr.subCategory);
                if (mr.category) keysToSet.add(mr.category);
                if (mr.uploadFolderId) keysToSet.add(mr.uploadFolderId);
                if (mr.vesselName && mr.subCategory) {
                  keysToSet.add(`${mr.vesselName} > ${mr.subCategory}`.toLowerCase());
                  keysToSet.add(`${mr.vesselName} > ${mr.group} > ${mr.subCategory}`.toLowerCase());
                }
              }
            }

            for (const key of Array.from(keysToSet)) {
              const existing = uploadedFilesByFolder[key] || [];
              uploadedFilesByFolder[key] = [...existing.filter(f => f.name !== file.name), newUpload];
            }

            // Also update matching row in rows array
            const targetParts = (targetSubFolderPath || '').split('>').map(s => s.trim()).filter(Boolean);
            const targetLeaf = targetParts[targetParts.length - 1] || (currentFolderNode?.name || 'Uploaded Files');
            const targetCategory = targetParts.length > 1 ? targetParts[targetParts.length - 2] : targetLeaf;
            const refRow = matchingRows[0] || baseRows.find(r => r.vesselName?.toLowerCase() === vLower) || null;

            const targetGroupKey = refRow
              ? `${refRow.vesselName}||${refRow.group}||${targetCategory}||${targetLeaf}||${targetSubFolderPath}`
              : `bulk||${targetCategory}||${targetLeaf}||${targetSubFolderPath}`;
            const targetUploadFolderId = result.folderId || refRow?.uploadFolderId || folderId;

            const effectiveRefRow: FlatRow = refRow || {
              srNo: String(baseRows.length + 1),
              vesselName: vesselName || '',
              group: (host.state.docMainFolder || 'Technical & Crewing'),
              category: targetCategory,
              subCategory: targetLeaf,
              subFolderPath: targetSubFolderPath,
              fileName: null,
              fileId: null,
              filePending: false,
              fileUploadedAt: undefined,
              groupKey: targetGroupKey,
              uploadFolderId: targetUploadFolderId,
            };

            let updatedRows = [...baseRows];

            // Ensure folder row exists for this targetSubFolderPath (fileName: null)
            const hasFolderRow = updatedRows.some(r =>
              (r.subFolderPath || '').toLowerCase() === targetNorm && !r.fileName
            );
            if (!hasFolderRow) {
              const nextSrBase = String(updatedRows.length + 1);
              const folderRow: FlatRow = {
                ...effectiveRefRow,
                srNo: nextSrBase,
                category: targetCategory,
                subCategory: targetLeaf,
                subFolderPath: targetSubFolderPath,
                fileName: null,
                fileId: null,
                filePending: false,
                fileUploadedAt: undefined,
                groupKey: targetGroupKey,
                uploadFolderId: targetUploadFolderId,
              };
              updatedRows.push(folderRow);
            }

            // Check if file row already exists for this file
            const existingFileIdx = updatedRows.findIndex(r =>
              ((r.subFolderPath || '').toLowerCase() === targetNorm || (r.groupKey || '').toLowerCase() === targetGroupKey.toLowerCase()) &&
              r.fileName?.toLowerCase() === file.name.toLowerCase()
            );

            if (existingFileIdx !== -1) {
              updatedRows = updatedRows.map((r, idx) =>
                idx === existingFileIdx
                  ? { ...r, fileName: file.name, fileId, filePending: result.statusPending, fileUploadedAt: newUpload.uploadedAt }
                  : r
              );
            } else {
              const nextFileSr = `${updatedRows.length + 1}.1`;
              const fileRow: FlatRow = {
                ...effectiveRefRow,
                srNo: nextFileSr,
                category: targetCategory,
                subCategory: targetLeaf,
                subFolderPath: targetSubFolderPath,
                groupKey: targetGroupKey,
                uploadFolderId: targetUploadFolderId,
                fileName: file.name,
                fileId,
                filePending: result.statusPending,
                fileUploadedAt: newUpload.uploadedAt,
              };
              updatedRows.push(fileRow);
            }

            // Optimistically update site folder cache for parent folder in SharePoint Sites view
            if (targetSiteId && targetDriveId && relativePath && relativePath.includes('/')) {
              const topFolderSeg = relativePath.split('/')[0];
              const pKeys = [
                currentFolderNode?.id ? `${targetSiteId}::${targetDriveId}::${currentFolderNode.id}` : '',
                folderId ? `${targetSiteId}::${targetDriveId}::${folderId}` : '',
              ].filter(Boolean);
              const folderItem = {
                id: result.folderId || `folder_${topFolderSeg}`,
                name: topFolderSeg,
                folder: { childCount: 1 },
                webUrl: '',
                lastModifiedDateTime: new Date().toISOString(),
              };
              pKeys.forEach(pKey => {
                const pCache = host._siteFolderItemsCache.get(pKey);
                if (pCache) {
                  if (!pCache.items.some((it: any) => (it.name || '').toLowerCase() === topFolderSeg.toLowerCase())) {
                    pCache.items = [folderItem, ...pCache.items];
                  }
                } else {
                  host._siteFolderItemsCache.set(pKey, { items: [folderItem], loading: false, parentPath: '' });
                }
              });
            }

            return { uploadedFilesByFolder, rows: updatedRows };
          });

          lastWebUrl = result.webUrl || lastWebUrl;
          updateProgress(progressId, {
            status: 'done',
            message: (relativePath ? `→ ${relativePath}` : `→ ${targetSubFolderPath || 'current folder'}`) + tagNote,
            webUrl: result.webUrl || null,
          });
        } catch (err: any) {
          host._clearUploadingFilePlaceholder(optimisticKeys, optimisticId);
          console.error(`[VesselDMS] Bulk upload failed for "${file.name}":`, err);
          updateProgress(progressId, {
            status: 'failed',
            message: err?.message || 'Upload failed',
          });
        }
      }

      setDone(true);
      setLastSuccessWebUrl(lastWebUrl);
      setAutoCloseSeconds(10);
      void host._syncScheduler?.triggerNow().catch(() => undefined);
      if (targetSiteId && targetDriveId) {
        const refreshTargetId = currentFolderNode?.id || folderId || 'root';
        void host._refreshSiteFolder(targetSiteId, targetDriveId, refreshTargetId);
        if (folderId && folderId !== refreshTargetId) {
          void host._refreshSiteFolder(targetSiteId, targetDriveId, folderId);
        }
      }
    })();
  }, [reviewing]); // eslint-disable-line react-hooks/exhaustive-deps

  // Derived counts
  const doneCount = progress.filter(p => p.status === 'done').length;
  const failedCount = progress.filter(p => p.status === 'failed').length;
  const totalCount = files.length;
  const pct = totalCount > 0 ? Math.round((doneCount / totalCount) * 100) : 0;
  const canClose = !started || done;

  const statusIcon = (s: BulkUploadProgress['status']): React.ReactElement =>
    s === 'done' ? <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} />
    : s === 'failed' ? <Icon iconName="ErrorBadge" aria-hidden="true" style={{ fontSize: 12 }} />
    : s === 'uploading' ? <Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 12 }} />
    : <Icon iconName="CircleRing" aria-hidden="true" style={{ fontSize: 12 }} />;
  const statusColor = (s: BulkUploadProgress['status']): string =>
    s === 'done' ? '#15803d' : s === 'failed' ? '#dc2626' : s === 'uploading' ? '#0284c7' : '#94a3b8';
  const statusBg = (s: BulkUploadProgress['status']): string =>
    s === 'done' ? '#dcfce7' : s === 'failed' ? '#fee2e2' : s === 'uploading' ? '#e0f2fe' : '#f1f5f9';

  return (
    <div
      role="presentation"
      onClick={() => { if (canClose) onClose(); }}
      style={{
        position: 'fixed', inset: 0, zIndex: 100001,
        background: 'rgba(15,23,42,0.62)', backdropFilter: 'blur(4px)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 24,
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="bulk-upload-title"
        onClick={e => e.stopPropagation()}
        style={{
          width: isMobile ? '95vw' : 580, maxWidth: '95vw', background: '#fff', borderRadius: 20,
          boxShadow: '0 30px 80px rgba(15,23,42,0.38)',
          overflow: 'hidden', border: '1px solid #d1fae5',
          display: 'flex', flexDirection: 'column', maxHeight: '90vh',
        }}
      >
        {/* ── Header ── */}
        <div style={{
          padding: '22px 28px 18px', flexShrink: 0,
          background: done ? (failedCount > 0 ? '#fffbeb' : '#ecfdf5') : started ? '#f0f9ff' : '#f8fafc',
          borderBottom: '1px solid #e2e8f0',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
            <div style={{
              width: 52, height: 52, borderRadius: 14, flexShrink: 0,
              background: done ? (failedCount > 0 ? '#fef3c7' : '#dcfce7') : '#dbeafe',
              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 26,
            }}>
              {done
                ? (failedCount > 0
                  ? <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 26 }} />
                  : <Icon iconName="Completed" aria-hidden="true" style={{ fontSize: 26 }} />)
                : started
                ? <Icon iconName="Cloud" aria-hidden="true" style={{ fontSize: 26 }} />
                : <Icon iconName="Upload" aria-hidden="true" style={{ fontSize: 26 }} />}
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div id="bulk-upload-title" style={{
                fontSize: 18, fontWeight: 800, lineHeight: 1.2,
                color: done ? (failedCount > 0 ? '#b45309' : '#059669') : '#0f172a',
              }}>
                {done
                  ? (failedCount > 0 ? `Uploaded with ${failedCount} Error${failedCount !== 1 ? 's' : ''}` : <><Icon iconName="Completed" aria-hidden="true" style={{ fontSize: 18 }} /> Folder Upload Completed Successfully!</>)
                  : started
                  ? `Uploading ${currentIndex + 1} of ${totalCount} file${totalCount !== 1 ? 's' : ''}…`
                  : `Ready to Upload ${totalCount} File${totalCount !== 1 ? 's' : ''}`}
              </div>
              <div style={{ fontSize: 12, color: '#64748b', marginTop: 3 }}>
                {done
                  ? `${doneCount} of ${totalCount} file${totalCount !== 1 ? 's' : ''} uploaded successfully${failedCount > 0 ? `, ${failedCount} failed` : ''} · ${elapsedSeconds}s total`
                  : started
                  ? `${doneCount} done · Elapsed: ${elapsedSeconds}s`
                  : `Destination: ${subFolderPath || 'current folder'}`}
              </div>
            </div>
            {canClose && (
              <button onClick={onClose} style={{
                background: 'none', border: 'none', cursor: 'pointer',
                minWidth: 44, minHeight: 44, fontSize: 22, color: '#94a3b8', padding: 4, lineHeight: 1,
              }}>×</button>
            )}
          </div>

          {/* Progress bar */}
          {started && (
            <div style={{ marginTop: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: '#64748b', marginBottom: 5 }}>
                <span>{doneCount} / {totalCount} files</span>
                <span>{pct}%</span>
              </div>
              <div style={{ height: 8, background: '#e2e8f0', borderRadius: 999, overflow: 'hidden' }}>
                <div style={{
                  height: '100%', borderRadius: 999, transition: 'width 0.3s ease',
                  background: done
                    ? (failedCount > 0 ? 'linear-gradient(90deg,#f59e0b,#d97706)' : 'linear-gradient(90deg,#10b981,#059669)')
                    : 'linear-gradient(90deg,#0284c7,#38bdf8)',
                  width: `${pct}%`,
                }} />
              </div>
            </div>
          )}
        </div>

        {/* ── Tag review (before upload starts) ── */}
        {reviewing && tagOptions && (
          <div style={{ padding: '14px 28px 4px', flexShrink: 0, borderBottom: '1px solid #f1f5f9' }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: '#0f172a', marginBottom: 2 }}>Tags</div>
            <div style={{ fontSize: 11, color: '#64748b', marginBottom: 10 }}>
              Pre-filled from the destination folder. Change a value to override it for every file in this upload.
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 10, marginBottom: 10 }}>
              {TAG_FIELDS.map(f => {
                const value = tags[f.key];
                const auto = autoTagsRef.current[f.key];
                const options = tagOptions[f.key];
                return (
                  <label key={f.key} style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 11, color: '#64748b' }}>
                    <span style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>{f.label}</span>
                      {value && (
                        <span style={{ fontSize: 10, fontWeight: 700, color: value === auto ? '#0369a1' : '#b45309' }}>
                          {value === auto ? 'from folder' : 'edited'}
                        </span>
                      )}
                    </span>
                    <select
                      value={value}
                      onChange={e => setTags(prev => ({ ...prev, [f.key]: e.target.value }))}
                      style={{
                        minHeight: 36, padding: '6px 8px', borderRadius: 8, border: '1px solid #e2e8f0',
                        background: '#fff', color: '#0f172a', fontSize: 13, fontFamily: 'inherit',
                      }}
                    >
                      <option value="">(none)</option>
                      {options.map(o => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </label>
                );
              })}
            </div>
            {tagError && (
              <div style={{ fontSize: 11, color: '#b45309', background: '#fffbeb', borderRadius: 8, padding: '6px 10px', marginBottom: 10 }}>
                Couldn't load the tag lists ({tagError}). You can still upload; files get the usual automatic tags.
              </div>
            )}
            {!tagsEdited && !tagError && (
              <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 10 }}>
                No changes: files are tagged the usual automatic way.
              </div>
            )}
          </div>
        )}

        {/* ── File List ── */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '10px 28px' }}>
          {progress.map(p => (
            <div key={p.id} style={{
              display: 'flex', alignItems: 'center', gap: 10,
              padding: '9px 0', borderBottom: '1px solid #f1f5f9',
            }}>
              <div style={{
                width: 26, height: 26, borderRadius: '50%', flexShrink: 0,
                background: statusBg(p.status),
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: statusColor(p.status), fontWeight: 800, fontSize: 12,
              }}>
                {statusIcon(p.status)}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{
                  fontWeight: 600, fontSize: 13, color: '#1e293b',
                  overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                }}>
                  {p.name}
                </div>
                {(p.message || p.relativePath) && (
                  <div style={{
                    fontSize: 11, marginTop: 1,
                    color: p.status === 'failed' ? '#dc2626' : '#64748b',
                    overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                  }}>
                    {p.message || p.relativePath}
                  </div>
                )}
              </div>
              <div style={{ fontSize: 11, color: '#94a3b8', flexShrink: 0 }}>{p.size}</div>
              {p.status === 'done' && p.webUrl ? (
                <a
                  href={p.webUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  title="Open this file in SharePoint"
                  style={{
                    fontSize: 11, fontWeight: 700, flexShrink: 0, minWidth: 68, textAlign: 'right',
                    color: statusColor(p.status), textDecoration: 'none',
                  }}
                >
                  <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 11 }} /> Done <Icon iconName="Link" aria-hidden="true" style={{ fontSize: 11 }} />
                </a>
              ) : (
                <div style={{
                  fontSize: 11, fontWeight: 700, flexShrink: 0, minWidth: 68, textAlign: 'right',
                  color: statusColor(p.status),
                }}>
                  {p.status === 'done' ? <><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 11 }} /> Done</>
                    : p.status === 'failed' ? <><Icon iconName="ErrorBadge" aria-hidden="true" style={{ fontSize: 11 }} /> Failed</>
                    : p.status === 'uploading' ? 'Uploading…'
                    : 'Queued'}
                </div>
              )}
            </div>
          ))}
        </div>

        {/* ── Footer ── */}
        <div style={{
          padding: '14px 28px 20px', borderTop: '1px solid #e2e8f0',
          background: done ? (failedCount > 0 ? '#fffdf5' : '#f0fdf4') : '#fafafa', flexShrink: 0,
          display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexDirection: isMobile ? 'column' : 'row',
        }}>
          <div style={{ fontSize: 13, color: done && failedCount > 0 ? '#b45309' : '#15803d', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 6 }}>
            {done && failedCount > 0 && <><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 13 }} /> Upload completed: {doneCount} succeeded, {failedCount} failed.</>}
            {done && failedCount === 0 && <><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> Folder upload completed successfully ({doneCount} file{doneCount !== 1 ? 's' : ''} uploaded to SharePoint Online).</>}
            {!done && started && 'Files are being uploaded to SharePoint Online…'}
            {!done && !started && 'Click Upload to begin.'}
          </div>
          {done && doneCount > 0 ? (
            <div style={{ display: 'flex', gap: 8, width: isMobile ? '100%' : 'auto', flexDirection: isMobile ? 'column' : 'row' }}>
              <button
                onClick={() => {
                  host._toggleAlertBell();
                  onClose();
                }}
                style={{
                  background: '#0078d4', color: '#fff', border: 'none', borderRadius: 10, minHeight: 44, padding: '10px 20px',
                  fontSize: 13, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
                }}
              >
                View Alerts
              </button>
              {lastSuccessWebUrl && (
                <a
                  href={lastSuccessWebUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{
                    background: '#f1f5f9', color: '#0f172a', border: '1px solid #e2e8f0', borderRadius: 10,
                    minHeight: 44, padding: '10px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                    whiteSpace: 'nowrap', textDecoration: 'none', display: 'flex', alignItems: 'center', gap: 6,
                  }}
                >
                  <Icon iconName="Link" aria-hidden="true" style={{ fontSize: 13 }} /> Open in SharePoint
                </a>
              )}
              <button
                onClick={onClose}
                style={{
                  background: 'linear-gradient(135deg,#10b981,#059669)', color: '#fff', border: 'none', borderRadius: 10, minHeight: 44, padding: '10px 20px',
                  fontSize: 13, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
                }}
              >
                Stay here{autoCloseSeconds !== null ? ` (${autoCloseSeconds}s)` : ''}
              </button>
            </div>
          ) : reviewing ? (
            <div style={{ display: 'flex', gap: 8, width: isMobile ? '100%' : 'auto', flexDirection: isMobile ? 'column' : 'row' }}>
              <button
                onClick={onClose}
                style={{
                  background: '#f1f5f9', color: '#0f172a', border: '1px solid #e2e8f0', borderRadius: 10,
                  minHeight: 44, padding: '10px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                }}
              >
                Cancel
              </button>
              <button
                onClick={() => setReviewing(false)}
                disabled={!tagOptions}
                style={{
                  background: tagOptions ? '#0078d4' : '#f1f5f9', color: tagOptions ? '#fff' : '#94a3b8',
                  border: 'none', borderRadius: 10, minHeight: 44, padding: '10px 28px',
                  fontSize: 14, fontWeight: 700, cursor: tagOptions ? 'pointer' : 'not-allowed',
                }}
              >
                {tagOptions ? `Upload ${totalCount} file${totalCount !== 1 ? 's' : ''}` : 'Loading tags…'}
              </button>
            </div>
          ) : (
            <button
              onClick={onClose}
              disabled={!canClose}
              style={{
                background: done ? 'linear-gradient(135deg,#10b981,#059669)' : '#f1f5f9',
                color: done ? '#fff' : '#94a3b8',
                border: 'none', borderRadius: 10, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '10px 28px',
                fontSize: 14, fontWeight: 700,
                cursor: canClose ? 'pointer' : 'not-allowed',
                boxShadow: done ? '0 3px 10px rgba(16,185,129,0.3)' : 'none',
                transition: 'all 0.2s ease',
              }}
            >
              {done
                ? `Close${autoCloseSeconds !== null ? ` (${autoCloseSeconds}s)` : ''}`
                : 'Uploading…'}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Utility: extract all File entries from a DataTransfer (supports dropped folders) ──

export async function extractFilesFromDataTransfer(
  dataTransfer: DataTransfer
): Promise<BulkUploadFile[]> {
  const result: BulkUploadFile[] = [];

  const readEntry = (entry: any, path: string): Promise<void> => {
    return new Promise<void>((resolve) => {
      if (entry.isFile) {
        entry.file((file: File) => {
          result.push({ file, relativePath: path ? `${path}/${file.name}` : file.name });
          resolve();
        }, resolve);
      } else if (entry.isDirectory) {
        const reader = entry.createReader();
        const readAll = (): void => {
          reader.readEntries(async (entries: any[]) => {
            if (!entries.length) { resolve(); return; }
            for (const sub of entries) {
              await readEntry(sub, path ? `${path}/${entry.name}` : entry.name);
            }
            readAll(); // Read more batches (Chrome only returns 100 at a time)
          }, resolve);
        };
        readAll();
      } else {
        resolve();
      }
    });
  };

  const items = Array.from(dataTransfer.items || []);
  for (const item of items) {
    const entry = item.webkitGetAsEntry?.();
    if (entry) {
      await readEntry(entry, '');
    }
  }

  return result;
}
