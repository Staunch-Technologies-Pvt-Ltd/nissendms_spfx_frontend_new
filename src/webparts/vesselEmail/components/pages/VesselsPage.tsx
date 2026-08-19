import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import {
  badge, GROUP_COLORS, DATASOURCE_TAGS_MAP, VESSEL_TYPES, cleanName, suggestTagFromFilename,
  INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES, INITIAL_MOCK_APPROVALS,
  INITIAL_MOCK_NOTIFICATIONS, INITIAL_MOCK_USERS,
} from '../constants';
import type {
  FlatRow, GroupedRow, VesselRecord,
} from '../types/rows';
import type { BentoEmailLog } from '../types/bento';
import type { AppView, ModalMode } from '../types/view';
import type {
  FormState, DocPreviewItem, DeletedNode, DocumentItem, TemplateItem,
  ApprovalItem, NotificationItem, UserItem, FolderAnomalyItem, NormalFolderRecord,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';

function getSpoVesselFolderUrl(siteUrlProp?: string, vesselName?: string): string {
  const fallbackSite = 'https://nissenkaiunsingapore.sharepoint.com';
  const site = (siteUrlProp && siteUrlProp !== '#') ? siteUrlProp : fallbackSite;
  try {
    const urlObj = new URL(site);
    const basePath = urlObj.pathname.replace(/\/$/, '');
    const folderPath = vesselName
      ? `${basePath}/Shared Documents/Vessels/Specific Vessels/${vesselName}`
      : `${basePath}/Shared Documents/Vessels`;
    return `${urlObj.origin}${basePath}/Shared Documents/Forms/AllItems.aspx?id=${encodeURIComponent(folderPath)}`;
  } catch {
    return '#';
  }
}

// ── Dismiss Confirm Dialog ────────────────────────────────────────────────────
// Shown when user clicks "Dismiss" on any anomaly item.
// Gives two options: "Keep it here" (close dialog) or "Move to Recycle Bin".
export function renderDismissConfirmDialog(host: VesselEmail): React.ReactElement | null {
  const anomaly = host.state.spoAnomalyDismissConfirm;
  if (!anomaly) return null;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Dismiss Confirmation"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: 16,
      }}
      onClick={e => { if (e.target === e.currentTarget) host.setState({ spoAnomalyDismissConfirm: null }); }}
    >
      <div style={{
        background: '#fff', borderRadius: 16, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
        padding: 32, width: '100%', maxWidth: 440, position: 'relative',
      }}>
        {/* Close */}
        <button
          onClick={() => host.setState({ spoAnomalyDismissConfirm: null })}
          style={{ position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Cancel"
        >✕</button>

        {/* Icon + title */}
        <div style={{ textAlign: 'center', marginBottom: 20 }}>
          <div style={{ fontSize: 44, marginBottom: 10 }}>🗑️</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 17, fontWeight: 700, color: '#0f172a' }}>
            What would you like to do?
          </h3>
          <p style={{ margin: 0, fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>
            The {anomaly.item_type} <strong>"{anomaly.name}"</strong> was found outside the expected SharePoint structure.
          </p>
        </div>

        {/* Path */}
        <div style={{
          background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8,
          padding: '8px 12px', marginBottom: 22, fontSize: 11, color: '#475569',
          fontFamily: 'monospace', wordBreak: 'break-all',
        }}>
          📂 {anomaly.spo_path}
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {/* Keep it here */}
          <button
            onClick={() => host.setState({ spoAnomalyDismissConfirm: null })}
            style={{
              background: '#f8fafc', color: '#334155',
              border: '2px solid #e2e8f0', borderRadius: 10, padding: '14px 18px',
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
              transition: 'border-color 0.15s',
            }}
          >
            <span style={{ fontSize: 26 }}>📌</span>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2, color: '#0f172a' }}>Keep it here</div>
              <div style={{ fontSize: 12, color: '#64748b' }}>
                Leave this item listed in the warning section for now. You can classify or dismiss it later.
              </div>
            </div>
          </button>

          {/* Move to Recycle Bin */}
          <button
            onClick={() => host._moveAnomalyToRecycleBin(anomaly)}
            style={{
              background: 'linear-gradient(135deg, #ef4444, #dc2626)',
              color: '#fff', border: 'none', borderRadius: 10, padding: '14px 18px',
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
              transition: 'opacity 0.15s',
            }}
          >
            <span style={{ fontSize: 26 }}>🗑️</span>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2 }}>Move to Recycle Bin</div>
              <div style={{ fontSize: 12, opacity: 0.9 }}>
                Dismiss this warning and move the item record to the Recycle Bin. You can restore it from there.
              </div>
            </div>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Classify Dialog ──────────────────────────────────────────────────────────
export function renderClassifyDialog(host: VesselEmail): React.ReactElement | null {
  const dlg = host.state.spoClassifyDialog;
  if (!dlg) return null;

  return <ClassifyModalContent host={host} dlg={dlg} />;
}

function ClassifyModalContent({ host, dlg }: { host: VesselEmail; dlg: any }): React.ReactElement {

  const { anomaly, provisioning, done, doneNormal, alreadyExisted, error } = dlg;
  const [elapsed, setElapsed] = React.useState(0);
  const [autoClose, setAutoClose] = React.useState(8);

  React.useEffect(() => {
    if (!provisioning) {
      setElapsed(0);
      return;
    }
    const timer = setInterval(() => setElapsed(e => e + 1), 1000);
    return () => clearInterval(timer);
  }, [provisioning]);

  React.useEffect(() => {
    if (!done && !doneNormal) { setAutoClose(8); return; }
    const timer = setInterval(() => setAutoClose(s => {
      if (s <= 1) { clearInterval(timer); handleClose(); return 0; }
      return s - 1;
    }), 1000);
    return () => clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [done, doneNormal]);
  const formatTimer = (sec: number): string => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${m < 10 ? '0' + m : m}:${s < 10 ? '0' + s : s}`;
  };

  const handleVessel = async (): Promise<void> => {
    host.setState({ spoClassifyDialog: { ...dlg, provisioning: true, error: null } });
    try {
      // 1. Create vessel record via API
      const base = host._base();
      const newVessel = await host._fetchJson(`${base}/api/vessels`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: anomaly.name, imo: null, shipyard: 'Auto-Discovered', vessel_type: 'Bulk Carrier' }),
      });

      const newRecord: VesselRecord = {
        id: newVessel?.id || `v_${Date.now()}`,
        name: cleanName(anomaly.name),
        imo: newVessel?.imo || '—',
        status: 'Active',
        image_url: pickRandomVesselImage('Bulk Carrier'),
      };

      // 2. Add vessel directly to top of state grid so it appears at the top of the vessel list immediately
      host.setState(prev => ({
        vessels: [newRecord, ...prev.vessels.filter(v => v.name.toLowerCase() !== newRecord.name.toLowerCase())],
      }));

      // 3. Provision SPO DMS folder structure
      await host._provisionVesselFolders(anomaly.name, newRecord.id);

      // 4. Dismiss the anomaly
      host._dismissAnomaly(anomaly.id);

      // 5. Transition to success screen!
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: true, error: null } });
    } catch (e: any) {
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, error: e?.message || 'Failed to provision vessel.' } });
    }
  };

  const handleNormal = async (): Promise<void> => {
    host.setState({ spoClassifyDialog: { ...dlg, provisioning: true, error: null } });
    try {
      const result = await host._saveNormalFolder(anomaly);
      // Dismiss anomaly warning
      host._dismissAnomaly(anomaly.id);
      // Refresh normal folder list in state
      host._loadNormalFolders();
      // Show success screen
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, doneNormal: true, alreadyExisted: result?.already_existed === true, error: null } });
    } catch (e: any) {
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, doneNormal: false, error: e?.message || 'Failed to save folder.' } });
    }
  };

  const handleClose = (): void => {
    if (done || doneNormal) {
      host._loadData();
    }
    host.setState({ spoClassifyDialog: null });
  };

  const renderBody = (): React.ReactElement => {
    if (doneNormal) {
      return (
        <div style={{ textAlign: 'center', padding: '12px 0' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>📁</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0f172a' }}>
            {alreadyExisted ? 'Folder Already Registered' : 'Folder Confirmed as Normal Folder!'}
          </h3>
          <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            <strong>"{anomaly.name}"</strong> {alreadyExisted
              ? 'was already listed as a Normal Folder in the system.'
              : 'has been saved and will appear in the Normal Folders section below the vessel list.'}
          </p>
          <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: '#0369a1', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>📋 Saved to Normal Folders</div>
            <div>• Anomaly warning dismissed</div>
            <div>• Folder listed in "Normal Folders" section</div>
            <div>• Record stored in database</div>
          </div>
          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 12 }}>
            Auto-closing in {autoClose}s…
          </div>
          <button
            onClick={handleClose}
            style={{
              background: 'linear-gradient(135deg, #475569, #334155)',
              color: '#fff', border: 'none', borderRadius: 10,
              padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(71,85,105,0.3)',
            }}
          >
            Done / View Vessels
          </button>
        </div>
      );
    }

    if (provisioning) {
      return (
        <div style={{ textAlign: 'center', padding: '8px 0' }}>
          <div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' }}>
            Provisioning DMS Folder Structure…
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            Creating standard DMS folder hierarchy for <strong>"{anomaly.name}"</strong> in SharePoint Online. Please wait.
          </p>

          {/* Live Timer badge */}
          <div style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
            borderRadius: 20, padding: '8px 18px', fontSize: 14, fontWeight: 700, marginBottom: 20,
          }}>
            <span style={{ fontSize: 16 }}>⏱️</span>
            <span>Elapsed Time: {formatTimer(elapsed)}</span>
          </div>

          {/* Animated Progress Bar */}
          <div style={{ background: '#e2e8f0', borderRadius: 10, height: 8, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{
              background: 'linear-gradient(90deg, #0ea5e9, #0284c7, #38bdf8)',
              height: '100%', width: `${Math.min(96, 15 + elapsed * 12)}%`,
              transition: 'width 0.8s ease-out', borderRadius: 10,
            }} />
          </div>

    {/* Step Progress Checklist */}
          <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 }}>
            <div style={{ color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>✓</span><span>Vessel record confirmed in database</span>
            </div>
            <div style={{ color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>⏳</span><span>Creating SharePoint DMS folder tree (Technical &amp; Crewing, Month End, Certificates)…</span>
            </div>
            <div style={{ color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>○</span><span>Linking category subfolders &amp; permissions</span>
            </div>
          </div>

          {/* Live folder-creation feed — one line per folder as Graph confirms it */}
          {!!(host.state.folderCreationResults && host.state.folderCreationResults.length) && (
            <div style={{
              background: '#0f172a', borderRadius: 10, padding: '10px 12px', marginTop: 12,
              textAlign: 'left', fontSize: 11, fontFamily: 'monospace', maxHeight: 140,
              overflowY: 'auto', color: '#cbd5e1',
            }}>
              {host.state.folderCreationResults.map((r, i) => (
                <div key={`${r.path}-${i}`} style={{ display: 'flex', gap: 6, padding: '2px 0' }}>
                  <span>
                    {r.status === 'failed' ? '✗' : r.status === 'existed' ? '↩' : '✓'}
                  </span>
                  <span style={{
                    color: r.status === 'failed' ? '#f87171' : r.status === 'existed' ? '#94a3b8' : '#4ade80',
                    whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                  }}>
                    {r.path}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }

     if (done) {
      const failedCount = (host.state.folderCreationResults || []).filter(r => r.status === 'failed').length;
      return (
        <div style={{ textAlign: 'center', padding: '12px 0' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>{failedCount ? '⚠️' : '🎉'}</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: failedCount ? '#b45309' : '#059669' }}>
            {failedCount ? 'Vessel Provisioned With Some Issues' : 'Vessel Successfully Provisioned & Classified!'}
          </h3>
          <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            <strong>"{anomaly.name}"</strong> has been registered in the DMS database{failedCount ? ', but ' + failedCount + ' folder' + (failedCount === 1 ? '' : 's') + ' could not be created — see the log above.' : ' and its full SharePoint DMS folder tree has been created.'}
          </p>

          <div style={{ background: failedCount ? '#fffbeb' : '#ecfdf5', border: `1px solid ${failedCount ? '#fde68a' : '#a7f3d0'}`, borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: failedCount ? '#92400e' : '#065f46', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>{failedCount ? '⚠️ Provisioning Finished' : '✅ Provisioning Complete'}</div>
            <div>• Registered vessel card added to main grid</div>
            <div>• SharePoint DMS department subfolders created{failedCount ? ` (${failedCount} failed)` : ''}</div>
            <div>• Unrecognised warning dismissed</div>
          </div>

          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 12 }}>
            Auto-closing in {autoClose}s…
          </div>

          <button
            onClick={handleClose}
            style={{
              background: 'linear-gradient(135deg, #10b981, #059669)',
              color: '#fff', border: 'none', borderRadius: 10,
              padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
            }}
          >
            Done / View Vessels
          </button>
        </div>
      );
    }

    return (
      <>
        <div style={{ fontSize: 42, textAlign: 'center', marginBottom: 12 }}>
          {anomaly.item_type === 'folder' ? '📁' : '📄'}
        </div>

        <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a', textAlign: 'center' }}>
          Unrecognised SharePoint {anomaly.item_type === 'folder' ? 'Folder' : 'File'} Detected
        </h3>
        <p style={{ margin: '0 0 20px', fontSize: 13, color: '#64748b', textAlign: 'center' }}>
          A {anomaly.item_type} named <strong>"{anomaly.name}"</strong> was found directly inside the vessel management area in SharePoint Online but is not registered in DMS.
        </p>

        {/* Path */}
        <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 14px', marginBottom: 20, fontSize: 11, color: '#475569', fontFamily: 'monospace', wordBreak: 'break-all' }}>
          📂 {anomaly.spo_path}
        </div>

        {error && (
          <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' }}>
            <span>⚠️</span><span>{error}</span>
          </div>
        )}

        {anomaly.item_type === 'folder' ? (
          <>
            <p style={{ margin: '0 0 16px', fontSize: 13, fontWeight: 600, color: '#0f172a', textAlign: 'center' }}>
              What is this folder?
            </p>
            <div style={{ display: 'flex', gap: 12, flexDirection: 'column' }}>
              {/* Vessel option */}
              <button
                onClick={() => handleVessel()}
                style={{
                  background: 'linear-gradient(135deg, #0ea5e9, #0284c7)',
                  color: '#fff', border: 'none', borderRadius: 12, padding: '16px 20px',
                  cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
                  transition: 'opacity 0.2s',
                }}
              >
                <span style={{ fontSize: 28 }}>🚢</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>
                    This is a Vessel
                  </div>
                  <div style={{ fontSize: 12, opacity: 0.9 }}>
                    Register it as a vessel and create its full DMS folder structure in SharePoint
                  </div>
                </div>
              </button>

              {/* Normal folder option */}
              <button
                onClick={handleNormal}
                style={{
                  background: '#f8fafc', color: '#334155',
                  border: '2px solid #e2e8f0', borderRadius: 12, padding: '16px 20px',
                  cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
                  transition: 'opacity 0.2s',
                }}
              >
                <span style={{ fontSize: 28 }}>📁</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2, color: '#0f172a' }}>
                    This is a Normal Folder
                  </div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>
                    Dismiss the warning and keep it listed as an unstructured folder
                  </div>
                </div>
              </button>
            </div>
          </>
        ) : (
          <div style={{ display: 'flex', justifyContent: 'center' }}>
            <button
              onClick={handleNormal}
              style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
            >
              Dismiss Warning
            </button>
          </div>
        )}
      </>
    );
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Classify SharePoint Item"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: 16,
      }}
     onClick={e => { if (e.target === e.currentTarget) handleClose(); }}
    >
      <div style={{
        background: '#fff', borderRadius: 20, padding: '32px 36px',
        width: 480, maxWidth: '92vw',
        boxShadow: '0 24px 64px rgba(0,0,0,0.28)', position: 'relative',
      }}>
        <button
          onClick={handleClose}
          style={{ position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Close (provisioning continues in the background)"
        >✕</button>
        {renderBody()}
      </div>
    </div>
  );
}


// ── Unrecognised Folders Section ─────────────────────────────────────────────


function renderUnrecognisedFolders(host: VesselEmail, items: FolderAnomalyItem[]): React.ReactElement {
  return (
    <div style={{ marginTop: 32, background: 'linear-gradient(135deg, #fffbeb 0%, #fff9e6 100%)', border: '2px solid #f59e0b', borderRadius: 16, padding: 24 }}>
      {/* Header */}
      <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <div>
          <h3 style={{ margin: '0 0 4px', fontSize: 16, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 8 }}>
            ⚠️ Folders Created Outside DMS Records
            <span style={{ background: '#f59e0b', color: '#fff', borderRadius: 20, padding: '1px 10px', fontSize: 12, fontWeight: 700 }}>{items.length}</span>
          </h3>
          <p style={{ margin: 0, fontSize: 12, color: '#a16207' }}>
            These folders were created directly in SharePoint Online at the vessel management level but are not registered in DMS.
            Please classify each one as a Vessel or a Normal Folder.
          </p>
        </div>
      </div>

      {/* Items */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {items.map(item => (
          <div
            key={item.id}
            style={{
              background: '#fff', borderRadius: 12, border: '1px solid #fde68a',
              padding: '14px 18px', display: 'flex', alignItems: 'center',
              justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
              boxShadow: '0 1px 4px rgba(245,158,11,0.08)',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 }}>
              <span style={{ fontSize: 28 }}>📁</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14, color: '#1f1f1f', marginBottom: 2 }}>{item.name}</div>
                <div style={{ fontSize: 11, color: '#78716c', fontFamily: 'monospace' }}>{item.spo_path}</div>
                {item.detected_at && (
                  <div style={{ fontSize: 10, color: '#a8a29e', marginTop: 2 }}>
                    Detected: {new Date(item.detected_at).toLocaleString()}
                  </div>
                )}
              </div>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              {/* Classify button */}
              <button
                onClick={() => host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } })}
                style={{
                  background: 'linear-gradient(135deg, #f59e0b, #d97706)',
                  color: '#fff', border: 'none', borderRadius: 8, padding: '7px 14px',
                  fontSize: 12, fontWeight: 700, cursor: 'pointer',
                  display: 'inline-flex', alignItems: 'center', gap: 5,
                }}
              >
                🔍 Classify
              </button>

              {/* Open in SPO */}
              <a
                href={getSpoVesselFolderUrl(host.props.siteUrl, item.name)}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
                  borderRadius: 8, padding: '7px 12px', fontSize: 12, fontWeight: 600,
                  textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 4,
                }}
              >
                ↗ Open
              </a>

              {/* Dismiss → opens confirm dialog */}
              <button
                onClick={() => host.setState({ spoAnomalyDismissConfirm: item })}
                style={{
                  background: '#f1f5f9', color: '#64748b', border: '1px solid #cbd5e1',
                  borderRadius: 8, padding: '7px 12px', fontSize: 12, cursor: 'pointer',
                }}
                title="Dismiss warning"
              >
                ✕ Dismiss
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Vessel-Level Uploaded Files Section ──────────────────────────────────────
function renderVesselLevelFiles(host: VesselEmail, items: FolderAnomalyItem[]): React.ReactElement {
  const formatSize = (bytes?: number): string => {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1048576).toFixed(1)} MB`;
  };

  return (
    <div style={{ marginTop: 24, background: 'linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%)', border: '2px solid #a855f7', borderRadius: 16, padding: 24 }}>
      {/* Header */}
      <div style={{ marginBottom: 16 }}>
        <h3 style={{ margin: '0 0 4px', fontSize: 16, fontWeight: 700, color: '#6b21a8', display: 'flex', alignItems: 'center', gap: 8 }}>
          📄 Files Uploaded Outside DMS Structure
          <span style={{ background: '#a855f7', color: '#fff', borderRadius: 20, padding: '1px 10px', fontSize: 12, fontWeight: 700 }}>{items.length}</span>
        </h3>
        <p style={{ margin: 0, fontSize: 12, color: '#7c3aed' }}>
          These files were uploaded directly to the vessel management area in SharePoint Online — not inside any vessel's DMS folder structure.
        </p>
      </div>

      {/* Table-like list */}
      <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e9d5ff', overflow: 'hidden' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f3e8ff', borderBottom: '1px solid #e9d5ff' }}>
              <th style={{ padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8', width: 36 }}></th>
              <th style={{ padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8' }}>File Name</th>
              <th style={{ padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8' }}>Location</th>
              <th style={{ padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8' }}>Detected</th>
              <th style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 700, color: '#6b21a8' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {items.map((item, idx) => {
              const ext = item.name.split('.').pop()?.toUpperCase() || '';
              const extColor: Record<string, string> = { PDF: '#ef4444', DOCX: '#3b82f6', XLSX: '#10b981', PPTX: '#f59e0b', JPG: '#ec4899', PNG: '#ec4899' };
              const color = extColor[ext] || '#64748b';

              return (
                <tr
                  key={item.id}
                  style={{ borderBottom: idx < items.length - 1 ? '1px solid #f3e8ff' : 'none', background: idx % 2 === 0 ? '#fff' : '#faf5ff' }}
                >
                  <td style={{ padding: '10px 14px' }}>
                    <span style={{ background: `${color}18`, color, borderRadius: 4, padding: '2px 5px', fontSize: 10, fontWeight: 700 }}>
                      {ext || 'FILE'}
                    </span>
                  </td>
                  <td style={{ padding: '10px 14px', fontWeight: 600, color: '#1e293b' }}>{item.name}</td>
                  <td style={{ padding: '10px 14px', color: '#64748b', fontFamily: 'monospace', fontSize: 11, maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {item.spo_path}
                  </td>
                  <td style={{ padding: '10px 14px', color: '#94a3b8', whiteSpace: 'nowrap' }}>
                    {item.detected_at ? new Date(item.detected_at).toLocaleDateString() : '—'}
                  </td>
                  <td style={{ padding: '10px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                    <div style={{ display: 'inline-flex', gap: 6 }}>
                      <a
                        href={getSpoVesselFolderUrl(host.props.siteUrl, item.vessel_name || undefined)}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, textDecoration: 'none' }}
                      >
                        ↗ Open
                      </a>
                      <button
                        onClick={() => host.setState({ spoAnomalyDismissConfirm: item })}
                        style={{ background: '#f1f5f9', color: '#64748b', border: '1px solid #cbd5e1', borderRadius: 6, padding: '4px 8px', fontSize: 11, cursor: 'pointer' }}
                        title="Dismiss"
                      >✕ Dismiss</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── Normal Folders Section ────────────────────────────────────────────────────
function renderNormalFoldersSection(host: VesselEmail): React.ReactElement | null {
  const folders: NormalFolderRecord[] = host.state.normalFolders || [];
  if (folders.length === 0) return null;

  return (
    <div style={{ marginTop: 36 }}>
      {/* Divider */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
        <div style={{ flex: 1, height: 1, background: 'linear-gradient(to right, #94a3b8, transparent)' }} />
        <span style={{ fontSize: 12, fontWeight: 700, color: '#475569', background: '#f1f5f9', padding: '4px 14px', borderRadius: 20, border: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
          📁 Normal Folders ({folders.length})
        </span>
        <div style={{ flex: 1, height: 1, background: 'linear-gradient(to left, #94a3b8, transparent)' }} />
      </div>

      <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 14, padding: '16px 20px' }}>
        <p style={{ margin: '0 0 14px', fontSize: 12, color: '#64748b', lineHeight: 1.5 }}>
          The following folders were found in SharePoint Online but confirmed as <strong>normal (non-vessel) folders</strong> by a user. They are listed here for reference only.
        </p>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 10 }}>
          {folders.map((f, idx) => (
            <div
              key={f.id ?? `nf-${idx}`}
              style={{
                background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '12px 14px',
                display: 'flex', alignItems: 'flex-start', gap: 10,
              }}
            >
              <span style={{ fontSize: 24, flexShrink: 0, marginTop: 2 }}>
                {f.item_type === 'folder' ? '📁' : '📄'}
              </span>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {f.name}
                </div>
                <div style={{ fontSize: 10, color: '#94a3b8', fontFamily: 'monospace', marginTop: 2, wordBreak: 'break-all' }}>
                  {f.spo_path}
                </div>
                {f.detected_at && (
                  <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 4 }}>
                    Saved: {new Date(f.detected_at).toLocaleDateString()}
                  </div>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Vessel Card Provision Dialog ──────────────────────────────────────────────
// Shown when user clicks the "📁 Provision" button on a vessel card.

function renderProvisionDialog(host: VesselEmail): React.ReactElement | null {
  if (!host.state.spoProvisionDialog) return null;
  return <ProvisionModalContent host={host} dlg={host.state.spoProvisionDialog} />;
}
/** Vertical auto-scrolling ticker: shows each folder the moment the backend reports it created. */
function FolderCreationTicker({ feed }: { feed: any[] }): React.ReactElement | null {
  const scrollRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [feed.length]);

  if (!feed.length) return null;

  return (
    <div
      ref={scrollRef}
      style={{
        marginTop: 12, background: '#0f172a', borderRadius: 12, padding: '10px 14px',
        textAlign: 'left', fontSize: 11.5, fontFamily: 'monospace', maxHeight: 130,
        overflowY: 'auto', scrollBehavior: 'smooth',
      }}
    >
      {feed.map((entry, i) => (
        <div
          key={`${entry.path}-${i}`}
          style={{
            display: 'flex', alignItems: 'center', gap: 8, padding: '3px 0',
            color: entry.status === 'failed' ? '#f87171' : entry.status === 'existed' ? '#94a3b8' : '#4ade80',
          }}
        >
          <span>{entry.status === 'failed' ? '✗' : entry.status === 'existed' ? '↩' : '✓'}</span>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {entry.path.split('/').slice(-1)[0]}
          </span>
        </div>
      ))}
    </div>
  );
}
function ProvisionModalContent({ host, dlg }: { host: VesselEmail; dlg: any }): React.ReactElement {
  const { vessel, provisioning, done, error } = dlg;
  const [elapsed, setElapsed] = React.useState(0);
  const [autoClose, setAutoClose] = React.useState(10);

  const isProvisioned = (
    host.state.provisionedVesselIds?.has(vessel.id) ||
    (host.state.rows && host.state.rows.some(r => r.vesselName && r.vesselName.toLowerCase() === vessel.name.toLowerCase())) ||
    vessel.is_provisioned === true
  );

  React.useEffect(() => {
    if (!provisioning) { setElapsed(0); return; }
    const timer = setInterval(() => setElapsed(e => e + 1), 1000);
    return () => clearInterval(timer);
  }, [provisioning]);

  React.useEffect(() => {
    if (!done) { setAutoClose(10); return; }
    const timer = setInterval(() => setAutoClose(s => {
      if (s <= 1) { clearInterval(timer); handleClose(); return 0; }
      return s - 1;
    }), 1000);
    return () => clearInterval(timer);
  }, [done]);

  const formatTimer = (sec: number): string => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${m < 10 ? '0' + m : m}:${s < 10 ? '0' + s : s}`;
  };

  const handleProvision = async (): Promise<void> => {
    host.setState({ spoProvisionDialog: { ...dlg, provisioning: true, done: false, error: null } });
    try {
      const { success, results } = await host._provisionVesselFolders(vessel.name, vessel.id);
      const failedCount = results.filter(r => r.status === 'failed').length;
      host.setState({
        spoProvisionDialog: {
          ...dlg,
          provisioning: false,
          done: true,
          error: success ? null : `${failedCount} folder${failedCount === 1 ? '' : 's'} could not be created — see details below.`,
        },
      });
    } catch (e: any) {
      host.setState({ spoProvisionDialog: { ...dlg, provisioning: false, done: false, error: e?.message || 'Failed to provision vessel folders.' } });
    }
  };

  const handleClose = (): void => {
    host.setState({ spoProvisionDialog: null, folderCreationResults: null });
  };

  const renderBody = (): React.ReactElement => {
    if (provisioning) {
      return (
        <div style={{ textAlign: 'center', padding: '8px 0' }}>
          <div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' }}>
            Provisioning DMS Folder Structure…
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            Creating standard DMS folder hierarchy for <strong>"{ vessel.name }"</strong> in SharePoint Online. Please wait.
          </p>

          {/* Live Timer badge */}
          <div style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
            borderRadius: 20, padding: '8px 18px', fontSize: 14, fontWeight: 700, marginBottom: 20,
          }}>
            <span style={{ fontSize: 16 }}>⏱️</span>
            <span>Elapsed Time: {formatTimer(elapsed)}</span>
          </div>

          {/* Animated Progress Bar */}
          <div style={{ background: '#e2e8f0', borderRadius: 10, height: 8, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{
              background: 'linear-gradient(90deg, #0ea5e9, #0284c7, #38bdf8)',
              height: '100%', width: `${Math.min(96, 15 + elapsed * 12)}%`,
              transition: 'width 0.8s ease-out', borderRadius: 10,
            }} />
          </div>

          {/* Step Progress Checklist */}
          <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 }}>
            <div style={{ color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>✓</span><span>Vessel record confirmed in database</span>
            </div>
            <div style={{ color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>⏳</span><span>Creating SharePoint DMS folder tree (Technical &amp; Crewing, Month End, Certificates)…</span>
            </div>
         <div style={{ color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>○</span><span>Linking category subfolders &amp; permissions</span>
            </div>
          </div>

                  {/* Live Vertical News Ticker */}
          <FolderCreationTicker feed={host.state.folderCreationResults || []} />
        </div>
      );
    }

   if (done) {
      const spoUrl = getSpoVesselFolderUrl(host.props.siteUrl, vessel.name);
      const failedFolders = (host.state.folderCreationResults || []).filter(r => r.status === 'failed');
      const hasFailures = failedFolders.length > 0;
      return (
        <div style={{ textAlign: 'center', padding: '12px 0' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>{hasFailures ? '⚠️' : '🎉'}</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: hasFailures ? '#b45309' : '#059669' }}>
            {hasFailures ? 'Vessel Folders Provisioned (with some issues)' : 'Vessel Folders Successfully Provisioned!'}
          </h3>
          <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            {hasFailures ? (
              <><strong>"{vessel.name}"</strong> has been provisioned, but {failedFolders.length} folder{failedFolders.length === 1 ? '' : 's'} could not be created.</>
            ) : (
              <><strong>"{vessel.name}"</strong> has been fully provisioned — all SharePoint DMS folder structures have been created.</>
            )}
          </p>
          {hasFailures ? (
            <div style={{ background: '#fffbeb', border: '1px solid #fde68a', borderRadius: 10, padding: '12px 16px', marginBottom: 16, fontSize: 12, color: '#92400e', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, marginBottom: 6 }}>⚠️ Folders that failed:</div>
              <div style={{ maxHeight: 100, overflowY: 'auto' }}>
                {failedFolders.map((r, i) => (
                  <div key={`${r.path}-${i}`} title={r.error}>• {r.path.split('/').pop()}</div>
                ))}
              </div>
              <div style={{ marginTop: 8, color: '#78716c' }}>Click "🔄 Re-Provision Folders" after closing this dialog to retry.</div>
            </div>
          ) : (
            <div style={{ background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '12px 16px', marginBottom: 16, fontSize: 12, color: '#065f46', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, marginBottom: 4 }}>✅ Provisioning Complete</div>
              <div>• SharePoint DMS department subfolders created</div>
              <div>• IMO: {vessel.imo || '—'}</div>
              <div>• Vessel is ready for document uploads</div>
            </div>
          )}
          {/* Auto-close countdown */}
          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 14 }}>
            Auto-closing in {autoClose}s…
          </div>
          <div style={{ background: '#e2e8f0', borderRadius: 10, height: 4, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{
              background: '#10b981', height: '100%', borderRadius: 10,
              width: `${(autoClose / 10) * 100}%`, transition: 'width 1s linear',
            }} />
          </div>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button
              onClick={handleClose}
              style={{
                background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1',
                borderRadius: 10, padding: '10px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
            >
              Close
            </button>
            {spoUrl && spoUrl !== '#' && (
              <a
                href={spoUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  background: 'linear-gradient(135deg, #10b981, #059669)',
                  color: '#fff', border: 'none', borderRadius: 10,
                  padding: '10px 22px', fontSize: 13, fontWeight: 700,
                  textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 6,
                  boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
                }}
              >
                📂 Open Folder in SharePoint ↗
              </a>
            )}
          </div>
        </div>
      );
    }

    if (!provisioning && !done && isProvisioned) {
      return (
        <div style={{ textAlign: 'center', padding: '8px 0' }}>
          <div style={{ fontSize: 48, marginBottom: 10 }}>✅</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 19, fontWeight: 700, color: '#15803d' }}>
            Folders Already Provisioned
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            The standard SharePoint DMS folder tree for <strong>"{vessel.name}"</strong> is already provisioned and registered.
          </p>

          <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 10, padding: '12px 16px', marginBottom: 20, fontSize: 12, color: '#166534', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>📋 Provisioned Folder Structure:</div>
            <div>• Technical &amp; Crewing — Active</div>
            <div>• Commercial &amp; Chartering — Active</div>
            <div>• Insurance — Active</div>
          </div>

          {error && (
            <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' }}>
              <span>⚠️</span><span>{error}</span>
            </div>
          )}

          <div style={{ display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button
              onClick={handleClose}
              style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
            >
              Close
            </button>
            {host.props.siteUrl && (
              <a
                href={getSpoVesselFolderUrl(host.props.siteUrl, vessel.name)}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  background: 'linear-gradient(135deg, #0ea5e9, #0284c7)', color: '#fff',
                  border: 'none', borderRadius: 8, padding: '10px 20px',
                  fontSize: 13, fontWeight: 700, textDecoration: 'none',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
              >
                Open in SharePoint ↗
              </a>
            )}
            <button
              onClick={handleProvision}
              style={{
                background: 'transparent', color: '#0284c7', border: '1px solid #0284c7',
                borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
            >
              🔄 Re-Provision Folders
            </button>
          </div>
        </div>
      );
    }

    return (
      <>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{ fontSize: 44, marginBottom: 10 }}>📁</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a' }}>
            Provision SharePoint Folders
          </h3>
          <p style={{ margin: '0 0 12px', fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>
            This will create the full DMS folder hierarchy for <strong>"{ vessel.name }"</strong> in SharePoint Online across all departments.
          </p>
        </div>

        <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 10, padding: '12px 16px', marginBottom: 20, fontSize: 12, color: '#0369a1', textAlign: 'left' }}>
          <div style={{ fontWeight: 700, marginBottom: 4 }}>📋 What will be created:</div>
          <div>• Technical &amp; Crewing — monthly sub-folders + categories</div>
          <div>• Commercial — contract and invoice folders</div>
          <div>• Insurance — certificate and policy folders</div>
        </div>

        {error && (
          <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' }}>
            <span>⚠️</span><span>{error}</span>
          </div>
        )}

        <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
          <button
            onClick={handleClose}
            style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
          >
            Cancel
          </button>
          <button
            onClick={handleProvision}
            style={{
              background: 'linear-gradient(135deg, #0ea5e9, #0284c7)',
              color: '#fff', border: 'none', borderRadius: 8,
              padding: '10px 26px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(2,132,199,0.3)',
            }}
          >
            📁 Start Provisioning
          </button>
        </div>
      </>
    );
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Provision Vessel Folders"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        backdropFilter: 'blur(4px)',
      }}
      onClick={e => { if (e.target === e.currentTarget) handleClose(); }}
    >
      <div style={{
        background: '#fff', borderRadius: 20, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
        padding: 32, width: '100%', maxWidth: 480, position: 'relative',
      }}>
        <button
          onClick={handleClose}
          style={{ position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Close (provisioning continues in the background)"
        >✕</button>
        {renderBody()}
      </div>
    </div>
  );
}


// ── File Alert Dialog ────────────────────────────────────────────────────────
// Shown when a file is uploaded directly to SPO under the Vessels tree.
export function renderFileAlertDialog(host: VesselEmail): React.ReactElement | null {
  const dlg = host.state.spoFileAlertDialog;
  if (!dlg) return null;
  return <FileAlertDialogContent host={host} dlg={dlg} />;
}

function FileAlertDialogContent({ host, dlg }: { host: VesselEmail; dlg: import('../types/ui').SpoFileAlertDialog }): React.ReactElement {
  const { fileName, spoPath, moving, moved, error } = dlg;

  // ── Cascading selection state ──
  const [selVessel, setSelVessel] = React.useState(dlg.vesselName || '');
  const [selMainFolder, setSelMainFolder] = React.useState('');
  const [selGroupKey, setSelGroupKey] = React.useState('');

  const vessels = host.state.vessels;
  const rows = host.state.rows;

  // All unique main-folder (group) names for the selected vessel
  const mainFolders = React.useMemo(() => {
    if (!selVessel) return [];
    const norm = selVessel.trim().toLowerCase();
    return Array.from(new Set(
      rows.filter(r => r.vesselName.trim().toLowerCase() === norm && r.canUpload)
           .map(r => r.group)
    )).sort();
  }, [selVessel, rows]);

  // Sub-folder options for the selected vessel + main folder
  const subFolderOptions = React.useMemo(() => {
    if (!selVessel || !selMainFolder) return [];
    const norm = selVessel.trim().toLowerCase();
    return Array.from(
      new Map(
        rows
          .filter(r => r.vesselName.trim().toLowerCase() === norm && r.group === selMainFolder && r.canUpload)
          .map(r => [r.groupKey, { label: r.subFolderPath, groupKey: r.groupKey, uploadFolderId: r.uploadFolderId, subFolderPath: r.subFolderPath }])
      ).values()
    );
  }, [selVessel, selMainFolder, rows]);

  // Reset downstream selections when parent changes
  const onVesselChange = (v: string): void => { setSelVessel(v); setSelMainFolder(''); setSelGroupKey(''); };
  const onMainFolderChange = (v: string): void => { setSelMainFolder(v); setSelGroupKey(''); };

  const canMove = !!(selVessel && selMainFolder && selGroupKey && !moving);

  const handleMove = async (): Promise<void> => {
    if (!canMove) return;
    const target = subFolderOptions.find(o => o.groupKey === selGroupKey);
    if (!target) return;
    host.setState({ spoFileAlertDialog: { ...dlg, moving: true, error: null } });
    try {
      const { graphClient, siteId, driveId } = host.props;
      if (!graphClient || !siteId || !driveId) throw new Error('SharePoint context not available.');

      // Resolve the live SPO folder ID for the target subfolder
      const liveFolderId = host._getLiveSharePointFolderId(target.subFolderPath);
      let targetId = liveFolderId || target.uploadFolderId;

      // If targetId is a path (contains '/'), resolve it via Graph
      if (!targetId || targetId.includes('/')) {
        const parts = target.subFolderPath.split('>').map((p: string) => p.trim()).filter(Boolean);
        const vName = parts[0] || selVessel;
        const rest = parts.slice(1).join('/');
        const candidatePaths = [
          `Vessels/Specific Vessels/${vName}/${rest}`,
          `Vessels/${vName}/${rest}`,
        ].filter(Boolean);
        for (const tryPath of candidatePaths) {
          try {
            const enc = tryPath.split('/').map((s: string) => encodeURIComponent(s)).join('/');
            const item: any = await graphClient.api(`/sites/${siteId}/drives/${driveId}/root:/${enc}?$select=id,folder`).get();
            if (item?.id && item?.folder) { targetId = item.id; break; }
          } catch { /* try next */ }
        }
      }

      if (!targetId || targetId.includes('/')) throw new Error('Could not resolve target folder in SharePoint.');

      // Verify the source file still exists at its current location
      try {
        const srcItem: any = await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${dlg.fileId}?$select=id,name,parentReference`).get();
        console.log('[VesselDMS] Move source file:', srcItem?.name, 'parentId:', srcItem?.parentReference?.id, '→ targetId:', targetId);
      } catch (verifyErr: any) {
        throw new Error(`Source file not found in SharePoint (${verifyErr?.code || verifyErr?.message}). It may have already been moved.`);
      }

      // Move the file via Graph PATCH
      let moveResult: any;
      try {
        moveResult = await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${dlg.fileId}`)
          .patch({ parentReference: { id: targetId } });
      } catch (patchErr: any) {
        const msg = patchErr?.message || patchErr?.code || String(patchErr);
        throw new Error(`Graph move failed: ${msg}`);
      }
      if (!moveResult?.id) {
        throw new Error('Graph did not confirm the file move (no item ID returned).');
      }

      // Optimistically add the file to the target folder in UI state
      host.setState(prev => {
        const newFile = { name: fileName, size: '—', date: 'Just now', pending: false, id: dlg.fileId };
        const updated = { ...prev.uploadedFilesByFolder };
        const keys = [targetId, target.groupKey, target.uploadFolderId, target.subFolderPath.trim().toLowerCase()].filter(Boolean) as string[];
        keys.forEach(k => {
          updated[k] = [...(updated[k] || []).filter((f: any) => f.name !== fileName), newFile];
        });
        return { spoFileAlertDialog: { ...dlg, moving: false, moved: true, error: null }, uploadedFilesByFolder: updated };
      });

      // Refresh the target folder from SPO using the confirmed live folder ID
      // Pass isConfirmedGraphFolderId=true so _refreshFolderFiles uses ID lookup, not path lookup
      void host._refreshFolderFiles(targetId, target.groupKey, true, true).catch(() => undefined);
      void host._syncScheduler?.triggerNow().catch(() => undefined);
    } catch (e: any) {
      host.setState({ spoFileAlertDialog: { ...dlg, moving: false, error: e?.message || 'Move failed.' } });
    }
  };

  const handleClose = (): void => { host.setState({ spoFileAlertDialog: null }); };

  const selectStyle: React.CSSProperties = {
    width: '100%', padding: '9px 12px', borderRadius: 8,
    border: '1px solid #bfdbfe', fontSize: 12, background: '#fff',
    outline: 'none', marginBottom: 10, boxSizing: 'border-box',
  };

  return (
    <div
      role="dialog" aria-modal="true" aria-label="File Uploaded Outside DMS Structure"
      style={{ position: 'fixed', inset: 0, zIndex: 100010, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}
      onClick={e => { if (e.target === e.currentTarget && !moving) handleClose(); }}
    >
      <div style={{ background: '#fff', borderRadius: 20, boxShadow: '0 24px 64px rgba(0,0,0,0.28)', padding: 32, width: '100%', maxWidth: 520, position: 'relative' }}>
        {!moving && (
          <button onClick={handleClose} style={{ position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }} title="Close">✕</button>
        )}

        {moved ? (
          <div style={{ textAlign: 'center', padding: '12px 0' }}>
            <div style={{ fontSize: 52, marginBottom: 12 }}>✅</div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' }}>File Moved Successfully!</h3>
            <p style={{ margin: '0 0 8px', fontSize: 13, color: '#475569' }}>
              <strong>"{fileName}"</strong> has been moved to the selected DMS folder.
            </p>
            {selGroupKey && (() => {
              const t = subFolderOptions.find(o => o.groupKey === selGroupKey);
              return t ? <p style={{ margin: '0 0 20px', fontSize: 12, color: '#64748b', fontFamily: 'monospace' }}>📂 {t.subFolderPath}</p> : null;
            })()}
            <button onClick={handleClose} style={{ background: 'linear-gradient(135deg,#10b981,#059669)', color: '#fff', border: 'none', borderRadius: 10, padding: '10px 28px', fontSize: 14, fontWeight: 700, cursor: 'pointer' }}>Done</button>
          </div>
        ) : (
          <>
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
              <div style={{ width: 48, height: 48, borderRadius: 12, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, flexShrink: 0 }}>⚠️</div>
              <div>
                <h3 style={{ margin: 0, fontSize: 17, fontWeight: 700, color: '#0f172a' }}>File Uploaded Outside DMS Structure</h3>
                <p style={{ margin: '3px 0 0', fontSize: 12, color: '#64748b' }}>Detected in SharePoint but not inside a standard DMS folder.</p>
              </div>
            </div>

            {/* File info */}
            <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: '10px 14px', marginBottom: 16 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                <span style={{ fontSize: 18 }}>📄</span>
                <span style={{ fontWeight: 700, fontSize: 13, color: '#0f172a', wordBreak: 'break-all' }}>{fileName}</span>
              </div>
              <div style={{ fontSize: 11, color: '#64748b', fontFamily: 'monospace', wordBreak: 'break-all' }}>📂 {spoPath}</div>
            </div>

            {error && (
              <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '8px 12px', marginBottom: 12, fontSize: 12, color: '#dc2626' }}>⚠️ {error}</div>
            )}

            <p style={{ margin: '0 0 12px', fontSize: 13, fontWeight: 600, color: '#0f172a' }}>What would you like to do?</p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {/* Keep as is */}
              <button
                onClick={handleClose}
                style={{ background: '#f8fafc', color: '#334155', border: '2px solid #e2e8f0', borderRadius: 10, padding: '12px 16px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left' }}
              >
                <span style={{ fontSize: 24 }}>📌</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a', marginBottom: 2 }}>Keep as is</div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>Leave the file where it is. The alert stays in the bell for reference.</div>
                </div>
              </button>

              {/* Move to DMS subfolder — cascading selects */}
              <div style={{ background: 'linear-gradient(135deg,#eff6ff,#dbeafe)', border: '2px solid #bfdbfe', borderRadius: 10, padding: '14px 16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
                  <span style={{ fontSize: 24 }}>📁</span>
                  <div style={{ fontWeight: 700, fontSize: 13, color: '#1d4ed8' }}>Move to DMS Subfolder</div>
                </div>

                {/* Step 1: Vessel */}
                <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#1d4ed8', marginBottom: 4, textTransform: 'uppercase' }}>1. Vessel</label>
                <select value={selVessel} onChange={e => onVesselChange(e.target.value)} style={selectStyle}>
                  <option value="">— Select vessel —</option>
                  {vessels.map(v => <option key={v.id} value={v.name}>{v.name}</option>)}
                </select>

                {/* Step 2: Main Folder / Category */}
                <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#1d4ed8', marginBottom: 4, textTransform: 'uppercase' }}>2. Main Folder</label>
                <select value={selMainFolder} onChange={e => onMainFolderChange(e.target.value)} disabled={!selVessel || mainFolders.length === 0} style={{ ...selectStyle, opacity: !selVessel ? 0.5 : 1 }}>
                  <option value="">— Select main folder —</option>
                  {mainFolders.map(mf => <option key={mf} value={mf}>{mf}</option>)}
                </select>

                {/* Step 3: Sub-folder */}
                <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#1d4ed8', marginBottom: 4, textTransform: 'uppercase' }}>3. Sub-folder</label>
                <select value={selGroupKey} onChange={e => setSelGroupKey(e.target.value)} disabled={!selMainFolder || subFolderOptions.length === 0} style={{ ...selectStyle, opacity: !selMainFolder ? 0.5 : 1 }}>
                  <option value="">— Select sub-folder —</option>
                  {subFolderOptions.map(o => <option key={o.groupKey} value={o.groupKey}>{o.label}</option>)}
                </select>

                <button
                  onClick={() => void handleMove()}
                  disabled={!canMove}
                  style={{
                    width: '100%',
                    background: canMove ? 'linear-gradient(135deg,#3b82f6,#1d4ed8)' : '#93c5fd',
                    color: '#fff', border: 'none', borderRadius: 8, padding: '10px',
                    fontSize: 13, fontWeight: 700,
                    cursor: canMove ? 'pointer' : 'not-allowed',
                  }}
                >
                  {moving ? '⏳ Moving…' : '📁 Move File'}
                </button>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ── Main Page Render ─────────────────────────────────────────────────────────

function renderVesselActionPicker(host: VesselEmail): React.ReactElement {
  const action = host.state.vesselActionPicker;
  const isEdit = action === 'edit';

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={isEdit ? 'Select vessel to edit' : 'Select vessel to delete'}
      style={{ position: 'fixed', inset: 0, zIndex: 100010, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}
      onClick={e => { if (e.target === e.currentTarget) host.setState({ vesselActionPicker: null }); }}
    >
      <div style={{ background: '#fff', borderRadius: 16, width: '100%', maxWidth: 560, maxHeight: '82vh', overflow: 'hidden', boxShadow: '0 24px 64px rgba(0,0,0,0.28)' }}>
        <div style={{ padding: '20px 24px 16px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0, fontSize: 18, color: '#0f172a' }}>{isEdit ? 'Select Vessel to Edit' : 'Select Vessel to Delete'}</h3>
            <p style={{ margin: '5px 0 0', fontSize: 12, color: '#64748b' }}>
              {isEdit ? 'Choose a vessel to update its details.' : 'Choose a vessel to move it and its folder contents to the Recycle Bin.'}
            </p>
          </div>
          <button onClick={() => host.setState({ vesselActionPicker: null })} style={{ border: 'none', background: 'none', color: '#64748b', fontSize: 20, cursor: 'pointer' }} title="Close">✕</button>
        </div>
        <div style={{ padding: 16, maxHeight: 'calc(82vh - 100px)', overflowY: 'auto' }}>
          {host.state.vessels.length === 0 ? (
            <div style={{ padding: 20, textAlign: 'center', color: '#64748b', fontSize: 13 }}>No vessels are available.</div>
          ) : host.state.vessels.map(vessel => (
            <button
              key={vessel.id}
              onClick={() => isEdit ? host._openEditVessel(vessel) : host._openDeleteVessel(vessel)}
              style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 12, padding: '12px 14px', marginBottom: 8, border: '1px solid #e2e8f0', borderRadius: 10, background: '#fff', textAlign: 'left', cursor: 'pointer' }}
            >
              <span style={{ width: 36, height: 36, borderRadius: 8, background: isEdit ? '#e0f2fe' : '#fef2f2', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0 }}>{isEdit ? '✏️' : '🗑️'}</span>
              <span style={{ minWidth: 0, flex: 1 }}>
                <span style={{ display: 'block', fontSize: 14, fontWeight: 700, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{vessel.name}</span>
                <span style={{ display: 'block', marginTop: 3, fontSize: 11, color: '#64748b' }}>{vessel.vessel_type || 'Vessel'}{vessel.imo ? ` | IMO ${vessel.imo}` : ''}</span>
              </span>
              <span style={{ color: isEdit ? '#0284c7' : '#dc2626', fontSize: 12, fontWeight: 700 }}>{isEdit ? 'Edit →' : 'Delete →'}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

export function renderVesselsPage(host: VesselEmail): React.ReactElement {
    const {
      vessels, vesselsSearch, vesselStatusFilter, vesselTypeFilter,
      modal, selectedVessel, folderProvisioningVesselId, folderCreationError,
      folderCreationResults, panelLoading, loading,
    } = host.state;

    const filtered = vessels.filter(v => {
      if (vesselStatusFilter !== 'all' && (v.status || 'Active') !== vesselStatusFilter) return false;
      if (vesselTypeFilter && vesselTypeFilter !== 'all' && (v.vessel_type || '') !== vesselTypeFilter) return false;
      if (vesselsSearch) {
        const q = vesselsSearch.toLowerCase();
        return v.name.toLowerCase().includes(q) ||
          (v.imo || '').includes(vesselsSearch) ||
          (v.vessel_type || '').toLowerCase().includes(q) ||
          (v.shipyard || '').toLowerCase().includes(q);
      }
      return true;
    });

    // All unique vessel types for filter dropdown
    const allTypes = Array.from(new Set(vessels.map(v => v.vessel_type).filter(Boolean))) as string[];

    const isLoading = panelLoading || (loading && vessels.length === 0);

    // Segregated anomalies
    const allAnomalies = host.state.folderAnomalies || [];
    const vesselLevelFolders = allAnomalies.filter(a => a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'folder');
    const vesselLevelFiles   = allAnomalies.filter(a => a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'file');

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>

        {/* ── Header ── */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24, flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 28, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 12 }}>
              Vessels
              {vessels.length > 0 && (
                <span style={{ background: 'linear-gradient(135deg,#0078d4,#005a9e)', color: '#fff', borderRadius: 20, padding: '3px 14px', fontSize: 14, fontWeight: 700 }}>
                  {vessels.length}
                </span>
              )}
              {(vesselLevelFolders.length + vesselLevelFiles.length) > 0 && (
                <span
                  onClick={() => host._toggleAlertBell()}
                  style={{ background: '#fef3c7', color: '#92400e', borderRadius: 20, padding: '3px 12px', fontSize: 12, fontWeight: 700, border: '1px solid #f59e0b', display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}
                  title="SPO items detected outside DMS structure — click to view in Alerts"
                >
                  ⚠️ {vesselLevelFolders.length + vesselLevelFiles.length} SPO item{vesselLevelFolders.length + vesselLevelFiles.length !== 1 ? 's' : ''} — view in Alerts →
                </span>
              )}
            </h2>
            <p style={{ margin: '6px 0 0', fontSize: 14, color: '#64748b' }}>
              Manage fleet vessels, provision SharePoint folders, and view documents.
            </p>
          </div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            <button
              onClick={() => host._goToView('vessels').catch(() => undefined)}
              title="Reload vessel list from database"
              style={{ background: '#fff', color: '#475569', border: '1px solid #e2e8f0', borderRadius: 10, padding: '10px 18px', fontSize: 14, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 1px 3px rgba(16,24,40,.08)' }}
            >
              🔄 Refresh
            </button>
            <button
              onClick={() => host._openVesselActionPicker('edit')}
              style={{ background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 10, padding: '10px 20px', fontSize: 14, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 8px rgba(2,132,199,.3)' }}
            >
              ✏️ Edit
            </button>
            <button
              onClick={() => host._openVesselActionPicker('delete')}
              style={{ background: 'linear-gradient(135deg, #ef4444, #dc2626)', color: '#fff', border: 'none', borderRadius: 10, padding: '10px 20px', fontSize: 14, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 8px rgba(220,38,38,.3)' }}
            >
              🗑 Delete
            </button>
            <button
              onClick={host._openCreate}
              style={{ background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 10, padding: '10px 20px', fontSize: 14, fontWeight: 700, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 8px rgba(16,185,129,.3)' }}
            >
              + New Vessel
            </button>
          </div>
        </div>

        {/* ── Search & Filter Bar ── */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20, flexWrap: 'wrap' }}>
          <div style={{ position: 'relative', flex: '1 1 220px', minWidth: 180 }}>
            <span style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 14 }}>🔍</span>
            <input
              type="text"
              placeholder="Search by name, IMO, type, shipyard…"
              value={vesselsSearch}
              onChange={e => host.setState({ vesselsSearch: e.target.value })}
              style={{ width: '100%', padding: '10px 14px 10px 38px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: '#fff' }}
            />
          </div>
          <select
            value={vesselStatusFilter}
            onChange={e => host.setState({ vesselStatusFilter: e.target.value })}
            style={{ padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none', minWidth: 130 }}
          >
            <option value="all">All Status</option>
            <option value="Active">Active</option>
            <option value="In Maintenance">In Maintenance</option>
            <option value="Inactive">Inactive</option>
          </select>
          <select
            value={vesselTypeFilter || 'all'}
            onChange={e => host.setState({ vesselTypeFilter: e.target.value })}
            style={{ padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none', minWidth: 140 }}
          >
            <option value="all">All Types</option>
            {allTypes.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          {(vesselsSearch || vesselStatusFilter !== 'all' || (vesselTypeFilter && vesselTypeFilter !== 'all')) && (
            <button
              onClick={() => host.setState({ vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all' })}
              style={{ background: 'transparent', color: '#64748b', border: '1px solid #cbd5e1', borderRadius: 8, padding: '9px 14px', fontSize: 13, cursor: 'pointer' }}
            >
              ✕ Clear
            </button>
          )}
          <span style={{ marginLeft: 'auto', fontSize: 12, color: '#94a3b8', fontWeight: 500 }}>
            {filtered.length} vessel{filtered.length !== 1 ? 's' : ''}
          </span>
        </div>

        {/* ── Folder creation status banners ── */}
        {folderCreationError && (
          <div style={{ marginBottom: 12, background: '#fde7e9', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#a4262c', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>⚠ {folderCreationError}</span>
            <button onClick={() => host.setState({ folderCreationError: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#a4262c', fontWeight: 700 }}>✕</button>
          </div>
        )}
        {folderCreationResults && !folderCreationError && (
          <div style={{ marginBottom: 12, background: '#dff6dd', border: '1px solid #86efac', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#107c10', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
              <span>✅ SharePoint folders provisioned — {folderCreationResults.filter(r => r.status === 'created').length} created, {folderCreationResults.filter(r => r.status === 'existed').length} already existed.</span>
              {host.props.siteUrl && (
                <a
                  href={getSpoVesselFolderUrl(host.props.siteUrl)}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ color: '#0078d4', fontWeight: 600, textDecoration: 'underline', display: 'inline-flex', alignItems: 'center', gap: 3 }}
                >
                  Open SharePoint Folder ↗
                </a>
              )}
            </div>
            <button onClick={() => host.setState({ folderCreationResults: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#107c10', fontWeight: 700 }}>✕</button>
          </div>
        )}

        {/* ── Loading State ── */}
        {isLoading ? (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0' }}>
            <div style={{
              width: 40, height: 40, border: '3px solid #e2e8f0',
              borderTopColor: '#0078d4', borderRadius: '50%',
              animation: 'spin 0.8s linear infinite',
            }} />
            <p style={{ margin: 0, fontSize: 14, color: '#64748b', fontWeight: 500 }}>Loading vessels from database…</p>
            <p style={{ margin: 0, fontSize: 12, color: '#94a3b8' }}>Fetching vessel records and folder structure</p>
          </div>
        ) : filtered.length === 0 ? (
          /* ── Empty State ── */
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: '#fff', borderRadius: 12, border: '2px dashed #e2e8f0' }}>
            <span style={{ fontSize: 48 }}>🚢</span>
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>
                {vessels.length === 0 ? 'No vessels yet' : 'No vessels match your filter'}
              </p>
              <p style={{ margin: '6px 0 0', fontSize: 13, color: '#64748b' }}>
                {vessels.length === 0
                  ? 'Create your first vessel to provision its SharePoint folder structure.'
                  : 'Try clearing the search or filters.'}
              </p>
            </div>
            {vessels.length === 0 && (
              <button
                onClick={host._openCreate}
                style={{ background: '#10b981', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}
              >
                + Create First Vessel
              </button>
            )}
          </div>
        ) : (
          /* ── Vessel Cards Grid ── */
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 20 }}>
            {filtered.map(vessel => {
              const isSelected = selectedVessel?.id === vessel.id;
              const status = vessel.status || 'Active';
              const statusColor = status === 'Active' ? '#10b981' : status === 'In Maintenance' ? '#f59e0b' : '#ef4444';
              const statusBg   = status === 'Active' ? '#ecfdf5' : status === 'In Maintenance' ? '#fffbeb' : '#fef2f2';
              const isProvisioning = folderProvisioningVesselId === vessel.id;
              const imgSrc = getVesselImageForId(vessel.id);

              return (
                <div
                  key={vessel.id}
                  onClick={() => host.setState({ selectedVessel: isSelected ? null : vessel })}
                  style={{
                    background: '#fff',
                    borderRadius: 16,
                    border: isSelected ? '2px solid #0078d4' : '1px solid #e2e8f0',
                    boxShadow: isSelected
                      ? '0 0 0 4px rgba(0,120,212,0.15), 0 4px 16px rgba(0,0,0,0.10)'
                      : '0 1px 3px rgba(16,24,40,.08)',
                    overflow: 'hidden',
                    cursor: 'pointer',
                    transition: 'all 0.18s ease',
                    display: 'flex',
                    flexDirection: 'column',
                  }}
                  onMouseEnter={e => { if (!isSelected) { e.currentTarget.style.boxShadow = '0 4px 12px rgba(16,24,40,.12)'; e.currentTarget.style.transform = 'translateY(-1px)'; } }}
                  onMouseLeave={e => { if (!isSelected) { e.currentTarget.style.boxShadow = '0 1px 3px rgba(16,24,40,.08)'; e.currentTarget.style.transform = 'translateY(0)'; } }}
                >
                  {/* Card image header */}
                  <div style={{ position: 'relative', height: 140, overflow: 'hidden', background: '#1e3a5f' }}>
                    <img
                      src={resolveImgUrl(imgSrc)}
                      alt={vessel.name}
                      style={{ width: '100%', height: '100%', objectFit: 'cover', opacity: 0.9 }}
                    />
                    <div style={{
                      position: 'absolute', inset: 0,
                      background: 'linear-gradient(to bottom, rgba(0,0,0,0.1) 0%, rgba(15,23,42,0.8) 100%)',
                    }} />
                    {/* Status badge */}
                    <span style={{
                      position: 'absolute', top: 12, right: 12,
                      background: statusBg, color: statusColor,
                      borderRadius: 20, padding: '3px 12px', fontSize: 12, fontWeight: 700,
                      border: `1px solid ${statusColor}40`,
                    }}>
                      {status}
                    </span>
                    {/* Selection check */}
                    {isSelected && (
                      <span style={{
                        position: 'absolute', top: 12, left: 12,
                        background: '#0078d4', color: '#fff',
                        borderRadius: '50%', width: 26, height: 26,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        fontSize: 14, fontWeight: 700, boxShadow: '0 2px 8px rgba(0,0,0,0.3)',
                      }}>✓</span>
                    )}
                    {/* Vessel name */}
                    <div style={{ position: 'absolute', bottom: 12, left: 14, right: 14 }}>
                      <p style={{ margin: 0, fontSize: 17, fontWeight: 700, color: '#fff', textShadow: '0 1px 4px rgba(0,0,0,0.6)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                        {vessel.name}
                      </p>
                    </div>
                  </div>

                  {/* Card body */}
                  <div style={{ padding: '16px 18px', flex: 1, display: 'flex', flexDirection: 'column', gap: 10 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px 16px' }}>
                      {[['IMO', vessel.imo || '—'], ['Type', vessel.vessel_type || '—'], ['Shipyard', vessel.shipyard || '—'], ['Hull No.', vessel.hull_number || '—']].map(([label, val]) => (
                        <div key={label}>
                          <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>{label}</span>
                          <p style={{ margin: '2px 0 0', fontSize: 13, color: '#1e293b', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{val}</p>
                        </div>
                      ))}
                    </div>

                    {/* Action buttons */}
                    <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                      <button
                        onClick={e => {
                          e.stopPropagation();
                          host.setState({ vesselFilter: vessel.name, docMainFolder: null, folderPathStack: [] });
                          void host._goToView('list');
                        }}
                        style={{
                          flex: 1, background: 'linear-gradient(135deg,#eff6ff,#dbeafe)', color: '#1d4ed8',
                          border: '1px solid #bfdbfe', borderRadius: 8, padding: '8px 10px',
                          fontSize: 12, fontWeight: 700, cursor: 'pointer',
                          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                        }}
                      >
                        📄 View Documents
                      </button>
                      {(() => {
                        const provisioned = (
                          host.state.provisionedVesselIds?.has(vessel.id) ||
                          (host.state.rows && host.state.rows.some(r => r.vesselName && r.vesselName.toLowerCase() === vessel.name.toLowerCase())) ||
                          vessel.is_provisioned === true
                        );
                        return (
                          <button
                            disabled={!!folderProvisioningVesselId && folderProvisioningVesselId !== vessel.id}
                            onClick={e => {
                              e.stopPropagation();
                              host.setState({ spoProvisionDialog: { vessel, provisioning: isProvisioning, done: false, error: null } });
                            }}
                            style={{
                              flex: 1,
                              border: provisioned ? '1px solid #bbf7d0' : '1px solid #e2e8f0',
                              borderRadius: 8, padding: '8px 10px',
                              fontSize: 12, fontWeight: 700,
                              cursor: (folderProvisioningVesselId && folderProvisioningVesselId !== vessel.id) ? 'not-allowed' : 'pointer',                              background: isProvisioning ? '#f0f9ff' : provisioned ? 'linear-gradient(135deg,#f0fdf4,#dcfce7)' : '#f8fafc',
                              color: isProvisioning ? '#0284c7' : provisioned ? '#15803d' : '#334155',
                              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                            }}
                          >
                            {isProvisioning ? (
                              <><span style={{ display: 'inline-block', width: 11, height: 11, border: '2px solid #0284c7', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} /> Creating…</>
                            ) : provisioned ? '✓ Provisioned' : '📁 Provision'}
                          </button>
                        );
                      })()}
                      <a
                        href={getSpoVesselFolderUrl(host.props.siteUrl, vessel.name)}
                        target="_blank"
                        rel="noopener noreferrer"
                        onClick={e => e.stopPropagation()}
                        title="Open in SharePoint"
                        style={{
                          width: 36, display: 'flex', alignItems: 'center', justifyContent: 'center',
                          border: '1px solid #e2e8f0', borderRadius: 8, background: '#f8fafc',
                          color: '#0078d4', fontSize: 14, textDecoration: 'none',
                        }}
                      >
                        ↗
                      </a>
                    </div>

                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* ── Normal Folders Section ── */}
        {renderNormalFoldersSection(host)}

        {/* Keyframe animation for spinners */}
        <style>{`
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        `}</style>

        {/* ── Classify Dialog ── */}
        {renderClassifyDialog(host)}

        {/* ── Dismiss Confirm Dialog ── */}
        {renderDismissConfirmDialog(host)}

        {/* ── Provision Dialog ── */}
        {renderProvisionDialog(host)}

        {/* Modals */}
        {modal === 'create' && host._renderVesselForm('create')}
        {modal === 'edit' && host._renderVesselForm('edit')}
        {host.state.vesselActionPicker && renderVesselActionPicker(host)}
        {modal === 'delete' && host._renderDeleteModal()}
      </div>
    );
}
