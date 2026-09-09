import * as React from 'react';
import type VesselEmail from './VesselEmail';
import type { FlatRow } from './types/rows';
import type { VesselSuggestionUploadEntry } from './types/ui';
import { isMobileWidth } from './responsive';

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
}

export interface BulkUploadModalProps {
  host: VesselEmail;
  files: BulkUploadFile[];
  /** SPO drive item ID of the target folder */
  folderId: string;
  subFolderPath: string;
  vesselName: string;
  currentFolderNode: { id: string; name: string } | null;
  onClose: () => void;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function fmtSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

// ── Component ────────────────────────────────────────────────────────────────

export function BulkUploadModal({
  host, files, folderId, subFolderPath, vesselName, currentFolderNode, onClose,
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
  const hasRunRef = React.useRef(false);

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
    if (hasRunRef.current) return;
    hasRunRef.current = true;
    setStarted(true);

    void (async () => {
      const newUploads: Array<{ name: string; size: string; date: string; id?: string; uploadedAt: number }> = [];
      const successFilesList: File[] = []; // tracks files that uploaded OK — used for vessel suggestions
      const successUploadEntries: VesselSuggestionUploadEntry[] = [];

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
            file
          );

          host._replaceUploadingFilePlaceholder(optimisticKeys, optimisticId, file, {
            id: result.fileId,
            pending: result.statusPending,
            date: result.statusPending ? 'Queued for approval' : 'Just now',
            uploadedAt: Date.now(),
          });

          // Queue file into unified OCR staging pipeline (fire-and-forget)
          host._triggerOcrStaging(
            result.fileId,
            file,
            result.folderId || folderId,
            targetSubFolderPath,
            vesselName,
            'folder'
          );

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
            let updatedRows = baseRows;
            const targetParts = (targetSubFolderPath || '').split('>').map(s => s.trim()).filter(Boolean);
            const targetLeaf = targetParts[targetParts.length - 1] || (currentFolderNode?.name || 'Uploaded Files');
            const targetCategory = targetParts.length > 1 ? targetParts[targetParts.length - 2] : targetLeaf;
            const refRow = matchingRows[0] || baseRows.find(r => r.vesselName?.toLowerCase() === vLower) || null;

            const targetGroupKey = refRow
              ? `${refRow.vesselName}||${refRow.group}||${targetCategory}||${targetLeaf}||${targetSubFolderPath}`
              : `bulk||${targetCategory}||${targetLeaf}||${targetSubFolderPath}`;
            const targetUploadFolderId = result.folderId || refRow?.uploadFolderId || folderId;

            const existingRowIdx = baseRows.findIndex(r =>
              ((r.subFolderPath || '').toLowerCase() === targetNorm || (r.groupKey || '').toLowerCase() === targetGroupKey.toLowerCase()) &&
              !r.fileName
            );

            if (existingRowIdx !== -1) {
              updatedRows = baseRows.map((r, idx) =>
                idx === existingRowIdx
                  ? { ...r, fileName: file.name, fileId, filePending: result.statusPending, fileUploadedAt: newUpload.uploadedAt }
                  : r
              );
            } else if (refRow) {
              const alreadyHasFolderRow = baseRows.some(r => (r.subFolderPath || '').toLowerCase() === targetNorm);
              const nextSrBase = String(baseRows.length + 1);
              const folderRow = alreadyHasFolderRow
                ? null
                : {
                    ...refRow,
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
              const newRow = {
                ...refRow,
                srNo: `${nextSrBase}.1`,
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
              updatedRows = [...baseRows, newRow];
              if (folderRow) updatedRows = [...updatedRows, folderRow];
            }

            return { uploadedFilesByFolder, rows: updatedRows };
          });

          updateProgress(progressId, {
            status: 'done',
            message: relativePath ? `→ ${relativePath}` : `→ ${targetSubFolderPath || 'current folder'}`,
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
      setAutoCloseSeconds(10);
      void host._syncScheduler?.triggerNow().catch(() => undefined);

      // ── Vessel Suggestions: trigger after bulk upload completes ──
      // successFilesList is built synchronously inside the loop above, so no closure issue
      const filesToAnalyse = successFilesList.length > 0 ? successFilesList : files.map(f => f.file);
      if (filesToAnalyse.length > 0) {
        // Small delay so the bulk-upload completion UI is visible before the suggestion wizard opens
        window.setTimeout(() => {
          host._openVesselSuggestions(filesToAnalyse, vesselName || undefined, successUploadEntries);
        }, 800);
      }
    })();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Derived counts
  const doneCount = progress.filter(p => p.status === 'done').length;
  const failedCount = progress.filter(p => p.status === 'failed').length;
  const totalCount = files.length;
  const pct = totalCount > 0 ? Math.round((doneCount / totalCount) * 100) : 0;
  const canClose = !started || done;

  const statusIcon = (s: BulkUploadProgress['status']): string =>
    s === 'done' ? '✓' : s === 'failed' ? '✗' : s === 'uploading' ? '⏳' : '○';
  const statusColor = (s: BulkUploadProgress['status']): string =>
    s === 'done' ? '#15803d' : s === 'failed' ? '#dc2626' : s === 'uploading' ? '#0284c7' : '#94a3b8';
  const statusBg = (s: BulkUploadProgress['status']): string =>
    s === 'done' ? '#dcfce7' : s === 'failed' ? '#fee2e2' : s === 'uploading' ? '#e0f2fe' : '#f1f5f9';

  return (
    <div
      role="presentation"
      onClick={() => { if (canClose) onClose(); }}
      style={{
        position: 'fixed', inset: 0, zIndex: 10100,
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
              {done ? (failedCount > 0 ? '⚠️' : '🎉') : started ? '☁️' : '📤'}
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div id="bulk-upload-title" style={{
                fontSize: 18, fontWeight: 800, lineHeight: 1.2,
                color: done ? (failedCount > 0 ? '#b45309' : '#059669') : '#0f172a',
              }}>
                {done
                  ? (failedCount > 0 ? `Uploaded with ${failedCount} Error${failedCount !== 1 ? 's' : ''}` : '🎉 Folder Upload Completed Successfully!')
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
              <div style={{
                fontSize: 11, fontWeight: 700, flexShrink: 0, minWidth: 68, textAlign: 'right',
                color: statusColor(p.status),
              }}>
                {p.status === 'done' ? '✓ Done'
                  : p.status === 'failed' ? '✗ Failed'
                  : p.status === 'uploading' ? 'Uploading…'
                  : 'Queued'}
              </div>
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
            {done && failedCount > 0 && `⚠️ Upload completed: ${doneCount} succeeded, ${failedCount} failed.`}
            {done && failedCount === 0 && `✅ Folder upload completed successfully (${doneCount} file${doneCount !== 1 ? 's' : ''} uploaded to SharePoint Online).`}
            {!done && started && 'Files are being uploaded to SharePoint Online…'}
            {!done && !started && 'Click Upload to begin.'}
          </div>
          {done && doneCount > 0 ? (
            <div style={{ display: 'flex', gap: 8, width: isMobile ? '100%' : 'auto', flexDirection: isMobile ? 'column' : 'row' }}>
              <button
                onClick={() => {
                  host.setState({ alertOpen: true });
                  host._fetchAlerts();
                  onClose();
                }}
                style={{
                  background: '#0078d4', color: '#fff', border: 'none', borderRadius: 10, minHeight: 44, padding: '10px 20px',
                  fontSize: 13, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
                }}
              >
                View Alerts
              </button>
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
