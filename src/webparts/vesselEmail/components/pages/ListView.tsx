import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { FlatRow, GroupedRow } from '../types/rows';
import type { ApprovalItem } from '../types/ui';

const PAGE_ROWS = 10;

export function renderListView(
  host: VesselEmail,
  filtered: FlatRow[],
  groupedList: GroupedRow[],
): React.ReactElement {
  const { docListPage, docListSort, docUploadRowKey, docUploadBusy, textFilter, vesselFilter, docGroupFilter, catFilter, listViewSelectedFiles } = host.state;

   let sorted = [...groupedList];
  if (docListSort === 'name_az') {
    sorted.sort((a, b) => a.category.localeCompare(b.category));
  } else {
    sorted = sorted
      .map((row, i) => ({
        row, i,
        ts: row.files.reduce((max, f) => Math.max(max, (f as any).uploadedAt || 0), 0),
      }))
      .sort((a, b) => (b.ts - a.ts) || (a.i - b.i))
      .map(x => x.row);
  }

  const totalPages = Math.max(1, Math.ceil(sorted.length / PAGE_ROWS));
  const safePage   = Math.min(docListPage, totalPages - 1);
  const pageRows   = sorted.slice(safePage * PAGE_ROWS, (safePage + 1) * PAGE_ROWS);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', overflowX: 'auto', width: '100%' }}>
        <table style={{ width: '100%', minWidth: 850, borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
              <th style={{ padding: '10px 10px', width: 44, textAlign: 'center' }}>SR.</th>
              <th style={{ padding: '10px 12px' }}>VESSEL NAME</th>
              <th style={{ padding: '10px 12px' }}>GROUP</th>
              <th style={{ padding: '10px 12px' }}>CATEGORY</th>
              <th style={{ padding: '10px 12px' }}>SUB-CATEGORY</th>
              <th style={{ padding: '10px 12px' }}>FOLDER PATH</th>
              <th style={{ padding: '10px 12px' }}>FILE NAME</th>
              <th style={{ padding: '10px 12px', textAlign: 'right' }}>ATTACHMENT</th>
            </tr>
          </thead>
          <tbody>
            {pageRows.length === 0 ? (
              <tr>
                <td colSpan={8} style={{ padding: '36px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                  No documents found.{' '}
                  {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || catFilter !== 'all'
                    ? 'Try clearing the filters.' : ''}
                </td>
              </tr>
            ) : pageRows.map((r, idx) => {
              const globalIdx  = safePage * PAGE_ROWS + idx + 1;
              const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
              const hasFiles    = r.files.length > 0;

              return (
                <tr key={`${r.groupKey}-${idx}`}
                  style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                  onMouseLeave={e => (e.currentTarget.style.background = '')}
                >
                  <td style={{ padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center' }}>{globalIdx}</td>
                  <td style={{ padding: '10px 12px', fontWeight: 700, color: '#0f172a' }}>{r.vesselName}</td>
                  <td style={{ padding: '10px 12px' }}>
                    <span style={{ display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600, background: '#eff6ff', color: '#2563eb' }}>
                      {r.group}
                    </span>
                  </td>
                  <td style={{ padding: '10px 12px', fontWeight: 600, color: '#334155' }}>{r.category}</td>
                  <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b' }}>{r.subCategory || r.category}</td>
                  <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11 }} title={r.subFolderPath}>{r.subFolderPath}</td>
                  <td style={{ padding: '10px 12px' }}>
                    {hasFiles ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                        {r.files.map(file => (
                          <div key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                            <input
                              type="checkbox"
                              checked={listViewSelectedFiles.has(file.id)}
                              onChange={() => {
                                host.setState(prev => {
                                  const next = new Set(prev.listViewSelectedFiles);
                                  if (next.has(file.id)) next.delete(file.id); else next.add(file.id);
                                  return { listViewSelectedFiles: next };
                                });
                              }}
                              style={{ width: 14, height: 14, accentColor: '#ef4444', cursor: 'pointer', flexShrink: 0 }}
                            />
                            <span style={{ fontSize: 14 }}>📄</span>
                            <span
                              onClick={() => {
                                if (file.id && !file.id.startsWith('file_')) {
                                  void host._openDocumentFile(file.id, file.name);
                                } else {
                                  alert(`File "${file.name}" is pending — it will be available after approval.`);
                                }
                              }}
                              style={{ color: '#0284c7', textDecoration: 'underline', fontWeight: 600, cursor: 'pointer' }}
                              title={file.id && /^\d+$/.test(file.id) ? `${file.name} (pending approval)` : `Click to open ${file.name}`}
                            >
                              {file.name}{file.id && /^\d+$/.test(file.id) ? ' ⏳' : ''}
                            </span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <span style={{ color: '#94a3b8', fontStyle: 'italic', fontSize: 11 }}>—</span>
                    )}
                  </td>
                  <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 }}>
                      <div style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
                      <span
                        title={hasFiles ? 'Attachment Available' : 'Attachment Required'}
                        style={{ color: hasFiles ? '#15803d' : '#64748b', fontSize: 11, fontWeight: 600, whiteSpace: 'nowrap' }}
                      >
                        {hasFiles ? '✅ Attached' : '⚪ Not Attached'}
                      </span>
                      {/* Row-level upload */}
                      <label style={{
                        background: isUploading ? '#f1f5f9' : '#fff', border: '1px solid #cbd5e1', borderRadius: 6,
                        padding: '3px 8px', fontSize: 11, fontWeight: 600,
                        color: isUploading ? '#94a3b8' : '#334155',
                        cursor: isUploading ? 'not-allowed' : 'pointer',
                        display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
                      }}>
                        <input
                          type="file"
                          style={{ display: 'none' }}
                          disabled={isUploading}
                          onChange={async e => {
                            const file = e.target.files?.[0];
                            if (!file) return;
                            host.setState({ docUploadRowKey: r.groupKey, docUploadBusy: true, docUploadMsg: null });
                            try {
                              const { fileId, statusPending, folderId, isGraphUpload } = await host._uploadFileToFolder(
                                r.uploadFolderId, r.subFolderPath, r.vesselName, file, r.monthDriven
                              );
                              const msg = statusPending
                                ? `"${file.name}" submitted for approval.`
                                : `"${file.name}" uploaded successfully!`;

                              if (statusPending) {
                                const newApproval: ApprovalItem = {
                                  id: fileId || `a_${Date.now()}`,
                                  documentName: file.name,
                                  vessel: r.vesselName,
                                  requestedBy: host.props.userDisplayName || host.props.userEmail || 'You',
                                  requestedOn: new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
                                  status: 'Pending',
                                };
                                host.setState(prev => ({ approvalsList: [newApproval, ...prev.approvalsList] }));
                              }

                              const newUpload = { name: file.name, size: `${(file.size / 1024).toFixed(1)} KB`, date: 'Just now', pending: statusPending, id: fileId, uploadedAt: Date.now() };
                              const liveFolderId = folderId || r.uploadFolderId;

                              host.setState(prev => {
                                const uploadedFilesByFolder = prev.uploadedFilesByFolder;
                                const baseRows = (prev.rows || []) as FlatRow[];
                                const normSub   = (r.subFolderPath || '').trim().toLowerCase();
                                const dedupeKey = (r.subFolderPath || r.groupKey).trim().toLowerCase();

                                const updatedByFolder: Record<string, any[]> = {
                                  ...uploadedFilesByFolder,
                                  [r.groupKey]: [...(uploadedFilesByFolder[r.groupKey] || []).filter((f: any) => f.name !== file.name), newUpload],
                                };
                                // Write to liveFolderId only if it's a real Graph drive item ID (isGraphUpload = true)
                                // Do NOT write to r.uploadFolderId — it may be a backend DB ID shared across rows
                                if (isGraphUpload === true && liveFolderId) {
                                  updatedByFolder[liveFolderId] = [...(uploadedFilesByFolder[liveFolderId] || []).filter((f: any) => f.name !== file.name), newUpload];
                                }
                                if (normSub) {
                                  updatedByFolder[normSub] = [...(uploadedFilesByFolder[normSub] || []).filter((f: any) => f.name !== file.name), newUpload];
                                }
                                if (dedupeKey && dedupeKey !== normSub) {
                                  updatedByFolder[dedupeKey] = [...(uploadedFilesByFolder[dedupeKey] || []).filter((f: any) => f.name !== file.name), newUpload];
                                }

                                const existingRowIdx = baseRows.findIndex((row: FlatRow) => row.groupKey === r.groupKey && !row.fileName);
                                let updatedRows: FlatRow[];
                                if (existingRowIdx !== -1) {
                                  updatedRows = baseRows.map((row: FlatRow, i: number) =>
                                    i === existingRowIdx
                                      ? { ...row, fileName: file.name, fileId: fileId || `file_${Date.now()}`, filePending: statusPending }
                                      : row
                                  );
                                } else {
                                  const refRow = baseRows.find((row: FlatRow) => row.groupKey === r.groupKey) || r;
                                  updatedRows = [...baseRows, { ...refRow, fileName: file.name, fileId: fileId || `file_${Date.now()}`, filePending: statusPending }];
                                }
                                return { docUploadBusy: false, docUploadRowKey: null, docUploadMsg: msg, uploadedFilesByFolder: updatedByFolder, rows: updatedRows };
                              }, () => {
                                const refreshFolderId = folderId || r.uploadFolderId;
                                // Use isGraphUpload flag instead of Boolean(folderId): folderId may be a
                                // backend DB ID (not a real Graph drive item ID) when the REST fallback was used.
                                const isRealGraphId = isGraphUpload === true;
                                if (refreshFolderId && !/^f\d+$/.test(refreshFolderId)) {
                                  void host._refreshFolderFiles(refreshFolderId, r.groupKey, true, isRealGraphId).catch(() => undefined);
                                }
                              });
                            } catch (err: any) {
                              host.setState({ docUploadBusy: false, docUploadRowKey: null, docUploadMsg: `Upload failed: ${err?.message || 'Error'}` });
                            }
                          }}
                        />
                        {isUploading ? '⏳...' : '↑ Upload'}
                      </label>
                      {(() => {
                        const rowSelectedCount = r.files.filter(f => listViewSelectedFiles.has(f.id)).length;
                        return (
                          <button
                            type="button"
                            disabled={rowSelectedCount === 0}
                            onClick={() => {
                              if (rowSelectedCount === 0) return;
                              const filesToDelete = r.files
                                .filter(f => listViewSelectedFiles.has(f.id))
                                .map(f => ({ id: f.id, name: f.name, folderId: r.uploadFolderId, folderPath: r.subFolderPath }));
                              host.setState({ listViewSelectedFiles: new Set() });
                              host._openFileDeleteDialog(filesToDelete);
                            }}
                            style={{
                              border: `1px solid ${rowSelectedCount > 0 ? '#fca5a5' : '#e2e8f0'}`,
                              background: rowSelectedCount > 0 ? '#fff5f5' : '#f8fafc',
                              borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600,
                              cursor: rowSelectedCount > 0 ? 'pointer' : 'not-allowed',
                              color: rowSelectedCount > 0 ? '#ef4444' : '#cbd5e1',
                              display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
                            }}
                            title={rowSelectedCount > 0 ? `Delete ${rowSelectedCount} selected file(s)` : 'Select files to delete'}
                          >
                            🗑{rowSelectedCount > 0 ? ` Delete (${rowSelectedCount})` : ' Delete'}
                          </button>
                        );
                      })()}
                      </div>
                      {/* Open in SharePoint */}
                      <button
                        type="button"
                        onClick={() => void host._openSharePointFolder(r)}
                        title="Open this folder in SharePoint"
                        aria-label={`Open ${r.subFolderPath} in SharePoint`}
                        style={{ width: 28, height: 27, padding: 0, borderRadius: 6, border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', cursor: 'pointer', fontSize: 16, fontWeight: 700, lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                      >↗</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Pagination footer */}
        <div style={{ padding: '10px 14px', color: '#64748b', fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
          <span>
            Showing {filtered.length === 0 ? 0 : safePage * PAGE_ROWS + 1}–{Math.min((safePage + 1) * PAGE_ROWS, filtered.length)} of {filtered.length} rows
          </span>
          <div style={{ display: 'flex', gap: 3, alignItems: 'center' }}>
            <button
              onClick={() => host.setState({ docListPage: Math.max(0, safePage - 1) })}
              disabled={safePage === 0}
              style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}
            >‹</button>
            {Array.from({ length: totalPages }, (_, i) => (
              <button key={i} onClick={() => host.setState({ docListPage: i })}
                style={{ border: i === safePage ? 'none' : '1px solid #cbd5e1', background: i === safePage ? '#0078d4' : '#fff', color: i === safePage ? '#fff' : '#334155', borderRadius: 4, padding: '3px 8px', fontSize: 11, fontWeight: i === safePage ? 700 : 400, cursor: 'pointer', minWidth: 26 }}
              >{i + 1}</button>
            ))}
            <button
              onClick={() => host.setState({ docListPage: Math.min(totalPages - 1, safePage + 1) })}
              disabled={safePage >= totalPages - 1}
              style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}
            >›</button>
          </div>
        </div>
      </div>
    </div>
  );
}
