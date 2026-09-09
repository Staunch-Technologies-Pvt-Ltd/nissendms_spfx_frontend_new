import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { FlatRow, GroupedRow } from '../types/rows';
import type { ApprovalItem } from '../types/ui';
import { isMobileWidth, isTabletWidth } from '../responsive';
import { resolveDetectedVesselForFile } from '../constants';

const PAGE_ROWS = 10;

export function renderListView(
  host: VesselEmail,
  filtered: FlatRow[],
  groupedList: GroupedRow[],
): React.ReactElement {
  const { docListPage, docListSort, docUploadRowKey, docUploadBusy, textFilter, vesselFilter, docGroupFilter, catFilter, listViewSelectedFiles, documentFilesLoading, vesselLoadingName, documentVesselsLoadingMore } = host.state;
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const isMobile = isMobileWidth(viewportWidth);
  const isTablet = isTabletWidth(viewportWidth);

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
        {isMobile ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: 10 }}>
            {pageRows.map((r, idx) => {
              const globalIdx = safePage * PAGE_ROWS + idx + 1;
              const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
              const hasFiles = r.files.length > 0;
              const rowSelectedCount = r.files.filter(f => listViewSelectedFiles.has(f.id)).length;
              return (
                <div key={`${r.groupKey}-${idx}`} style={{ border: '1px solid #e2e8f0', borderRadius: 10, padding: 12 }}>
                  <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 6 }}>#{globalIdx}</div>
                  <div style={{ marginTop: 4, fontSize: 12, color: '#334155' }}><strong>{r.group}</strong> • {r.category} • {r.subCategory || r.category}</div>
                  <div style={{ marginTop: 8, fontSize: 11, color: '#64748b', wordBreak: 'break-word' }}>{r.subFolderPath}</div>
                  <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 6 }}>
                    {hasFiles ? r.files.map(file => (
                      <label key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
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
                          style={{ width: 16, height: 16, accentColor: '#ef4444' }}
                        />
                        <span
                          onClick={() => {
                            if ((file as any).uploading) {
                              alert(`File "${file.name}" is still uploading. Please wait a moment.`);
                            } else if (file.id && !file.id.startsWith('file_')) {
                              void host._openDocumentFile(file.id, file.name);
                            } else {
                              alert(`File "${file.name}" is pending — it will be available after approval.`);
                            }
                          }}
                          style={{ fontSize: 12, color: '#0369a1', textDecoration: 'underline', cursor: 'pointer' }}
                        >
                          {file.name}{(file as any).uploading ? ' ⏳' : ''}
                        </span>
                      </label>
                    )) : <span style={{ color: '#94a3b8', fontStyle: 'italic', fontSize: 11 }}>No files</span>}
                  </div>
                  <div style={{ marginTop: 10, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                    <label style={{ minHeight: 44, background: isUploading ? '#f1f5f9' : '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px', fontSize: 12, fontWeight: 700, textAlign: 'center', cursor: isUploading ? 'not-allowed' : 'pointer' }}>
                      <input type="file" multiple style={{ display: 'none' }} disabled={isUploading} onChange={e => {
                        const filesList = Array.from(e.target.files || []);
                        if (filesList.length === 0) return;
                        e.target.value = '';
                        const bulkFiles = filesList.map(f => ({ file: f, relativePath: f.name }));
                        host._openBulkUpload(bulkFiles, r.uploadFolderId, r.subFolderPath, r.vesselName);
                      }} />
                      Upload
                    </label>
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
                      style={{ minHeight: 44, border: '1px solid #fca5a5', background: rowSelectedCount > 0 ? '#fff5f5' : '#f8fafc', borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: rowSelectedCount > 0 ? 'pointer' : 'not-allowed', color: rowSelectedCount > 0 ? '#ef4444' : '#cbd5e1' }}
                    >
                      Delete {rowSelectedCount > 0 ? `(${rowSelectedCount})` : ''}
                    </button>
                  </div>
                  <button
                    type="button"
                    onClick={() => void host._openSharePointFolder(r)}
                    style={{ marginTop: 8, minHeight: 44, width: '100%', border: '1px solid #bfdbfe', background: '#eff6ff', borderRadius: 8, color: '#1d4ed8', fontWeight: 700, fontSize: 12, cursor: 'pointer' }}
                  >
                    Open Folder in SharePoint
                  </button>
                </div>
              );
            })}
            {pageRows.length === 0 && (
              <div style={{ padding: '24px 12px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                No documents found.
              </div>
            )}
          </div>
        ) : (
        <table style={{ width: '100%', minWidth: isTablet ? 980 : 850, borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
              <th style={{ padding: '10px 10px', width: 44, textAlign: 'center' }}>SR.</th>
              <th style={{ padding: '10px 12px' }}>MAIN FOLDER</th>
              <th style={{ padding: '10px 12px' }}>DOCUMENT SECTION</th>
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
                <td colSpan={9} style={{ padding: '48px 16px', textAlign: 'center' }}>
                  {(host.state.loading || Boolean(vesselLoadingName) || (documentFilesLoading && filtered.length === 0) || documentVesselsLoadingMore) ? (
                    <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
                      <div style={{
                        width: 32, height: 32, border: '3px solid #e0f2fe',
                        borderTop: '3px solid #0284c7', borderRadius: '50%',
                        animation: 'spin 0.8s linear infinite',
                      }} />
                      <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a' }}>
                        {vesselLoadingName
                          ? `Loading documents and attachments for ${vesselLoadingName}...`
                          : (documentVesselsLoadingMore ? 'Loading more vessels...' : 'Loading vessel documents from SharePoint...')}
                      </div>
                      <div style={{ fontSize: 11, color: '#64748b', maxWidth: 360 }}>
                        Please wait while folder structures and live files are loaded.
                      </div>
                    </div>
                  ) : (
                    <span style={{ color: '#94a3b8', fontSize: 13 }}>
                      No documents found.{' '}
                      {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || catFilter !== 'all'
                        ? 'Try clearing the filters.' : ''}
                    </span>
                  )}
                </td>
              </tr>
            ) : pageRows.map((r, idx) => {
              const globalIdx  = safePage * PAGE_ROWS + idx + 1;
              const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
              const hasFiles    = r.files.length > 0;
              const pathParts = (r.subFolderPath || '').split('>').map(s => s.trim()).filter(Boolean);
              const documentSectionLabel = pathParts.length >= 3 ? pathParts[2] : 'Drawings and Manuals';
              const categoryLabel = r.category || (pathParts.length >= 5 ? pathParts[4] : (pathParts[3] || 'To be Classified'));

              return (
                <tr key={`${r.groupKey}-${idx}`}
                  style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                  onMouseLeave={e => (e.currentTarget.style.background = '')}
                >
                  <td style={{ padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center' }}>{globalIdx}</td>
                  <td style={{ padding: '10px 12px' }}>
                    <span style={{ display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600, background: '#eff6ff', color: '#2563eb' }}>
                      {r.group}
                    </span>
                  </td>
                  <td style={{ padding: '10px 12px', fontWeight: 600, color: '#334155' }}>{documentSectionLabel}</td>
                  <td style={{ padding: '10px 12px', fontWeight: 700, color: '#1e293b' }}>{categoryLabel}</td>
                  <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b' }}>{r.subCategory || r.category}</td>
                  <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11 }} title={r.subFolderPath}>{r.subFolderPath}</td>
                  <td style={{ padding: '10px 12px' }}>
                    {hasFiles ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                        {r.files.map(file => (
                          <div key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
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
                            <span style={{ fontSize: 14 }}>{(file as any).uploading ? '⏳' : '📄'}</span>
                            <span
                              onClick={() => {
                                if ((file as any).uploading) {
                                  alert(`File "${file.name}" is still uploading. Please wait a moment.`);
                                } else if (file.id && !file.id.startsWith('file_')) {
                                  void host._openDocumentFile(file.id, file.name);
                                } else {
                                  alert(`File "${file.name}" is pending — it will be available after approval.`);
                                }
                              }}
                              style={{ color: '#0284c7', textDecoration: 'underline', fontWeight: 600, cursor: 'pointer', flexGrow: 1 }}
                              title={(file as any).uploading
                                ? `${file.name} (uploading)`
                                : (file.id && /^\d+$/.test(file.id) ? `${file.name} (pending approval)` : `Click to open ${file.name}`)}
                            >
                              {file.name}{(file as any).uploading ? ' ⏳' : (file.id && /^\d+$/.test(file.id) ? ' ⏳' : '')}
                            </span>
                            {(() => {
                              const detectedVessel = resolveDetectedVesselForFile(file, host, r.vesselName);
                              const isUnidentified = (host?.state?.ocrUnidentifiedFiles || []).some(
                                n => (n || '').trim().toLowerCase() === file.name.trim().toLowerCase()
                              );

                              if (detectedVessel && !isUnidentified) {
                                return (
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      host._openVesselSuggestions([new File([], file.name)], detectedVessel);
                                    }}
                                    title={`View OCR Vessel Suggestion: ${detectedVessel}`}
                                    style={{
                                      background: 'rgba(59,130,246,0.1)', border: '1px solid rgba(59,130,246,0.3)',
                                      borderRadius: 6, padding: '1px 6px', fontSize: 10, color: '#0284c7',
                                      cursor: 'pointer', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 3,
                                      whiteSpace: 'nowrap', flexShrink: 0,
                                    }}
                                  >
                                    <span>✨</span> {detectedVessel}
                                  </button>
                                );
                              } else {
                                return (
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      host._openVesselSuggestions([new File([], file.name)], r.vesselName || '');
                                    }}
                                    title="Vessel name not detected by OCR — click to assign vessel"
                                    style={{
                                      background: '#fff7ed', border: '1px solid #fed7aa',
                                      borderRadius: 6, padding: '1px 6px', fontSize: 10, color: '#c2410c',
                                      cursor: 'pointer', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 3,
                                      whiteSpace: 'nowrap', flexShrink: 0,
                                    }}
                                  >
                                    <span>⚠️</span> Vessel: Not detected
                                  </button>
                                );
                              }
                            })()}
                            {/* OCR Re-classify button for existing files */}
                            {!((file as any).uploading) && file.id && !file.id.startsWith('file_') && !/^\d+$/.test(file.id) && (
                              <button
                                type="button"
                                title={`Run OCR & Re-classify "${file.name}" into correct folder`}
                                onClick={() => {
                                  host._goToView('templates');
                                  // Store pending OCR item in state so TemplatesPage picks it up
                                  setTimeout(() => {
                                    host.setState({ ocrPendingItemId: file.id, ocrPendingFilename: file.name });
                                  }, 100);
                                }}
                                style={{
                                  background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 6,
                                  padding: '2px 7px', fontSize: 10, fontWeight: 700, cursor: 'pointer',
                                  color: '#15803d', display: 'inline-flex', alignItems: 'center', gap: 3,
                                  whiteSpace: 'nowrap', flexShrink: 0,
                                }}
                              >
                                🔍 OCR
                              </button>
                            )}
                            {!((file as any).uploading) && file.id && !file.id.startsWith('file_') && !/^\d+$/.test(file.id) && (
                              <button
                                type="button"
                                title={`Archive "${file.name}"`}
                                onClick={() => void host._archiveDocumentFile(file.id, file.name, r.subFolderPath, r.group, r.vesselName)}
                                style={{ border: '1px solid #c4b5fd', background: '#f5f3ff', borderRadius: 6, padding: '2px 7px', fontSize: 10, fontWeight: 700, cursor: 'pointer', color: '#6d28d9', whiteSpace: 'nowrap' }}
                              >
                                📦 Archive
                              </button>
                            )}
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
                          multiple
                          style={{ display: 'none' }}
                          disabled={isUploading}
                          onChange={e => {
                            const filesList = Array.from(e.target.files || []);
                            if (filesList.length === 0) return;
                            e.target.value = '';
                            const bulkFiles = filesList.map(f => ({ file: f, relativePath: f.name }));
                            host._openBulkUpload(bulkFiles, r.uploadFolderId, r.subFolderPath, r.vesselName);
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
        )}

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
