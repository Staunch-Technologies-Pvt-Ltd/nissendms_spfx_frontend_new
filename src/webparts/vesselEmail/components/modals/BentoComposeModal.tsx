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
  ApprovalItem, NotificationItem, UserItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import { isMobileWidth } from '../responsive';

export function renderBentoComposeModal(host: VesselEmail): React.ReactElement | null {
    const {
      bentoComposeOpen, bentoComposeForm, bentoComposeBusy,
      bentoComposeMsg, bentoComposeErr, vessels, bentoLogs, bentoApprovedFiles,
    } = host.state;
    if (!bentoComposeOpen) return null;
    const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

    const FIXED_RECIPIENT = host.state.bentoConfigRecipient || host.props.userEmail || '';

    const setForm = (patch: Partial<typeof bentoComposeForm>, autoSubject?: boolean): void => {
      const merged = { ...bentoComposeForm, ...patch };
      // Auto-fill subject whenever vessel or tag changes
      if (autoSubject) {
        merged.subject_text = host._buildAutoSubject(
          merged.vessel_name,
          merged.datasource_tag,
          (merged as any).subject_suffix || 'Subject'
        );
      }
      host.setState({ bentoComposeForm: merged });
    };

    // Approved files available for the selected vessel
    const approvedFiles = host._getApprovedFilesForVesselTag(
      bentoComposeForm.vessel_name, bentoComposeForm.datasource_tag
    );
    const vLowerKey = bentoComposeForm.vessel_name.trim().toLowerCase();
    const isLoadingFiles = !!bentoComposeForm.vessel_name && !(vLowerKey in (bentoApprovedFiles || {}));

    const handleSend = async (): Promise<void> => {
      if (!bentoComposeForm.subject_text.trim()) {
        host.setState({ bentoComposeErr: 'Subject is required.' }); return;
      }
      host.setState({ bentoComposeBusy: true, bentoComposeErr: null, bentoComposeMsg: null });

      // Add a pending log entry immediately
      const tempId = Date.now();
      const pendingLog: BentoEmailLog = {
        id: tempId,
        datasource_tag_used: bentoComposeForm.datasource_tag,
        tag_label: DATASOURCE_TAGS_MAP[bentoComposeForm.datasource_tag],
        tag_was_valid: true,
        vessel_name: bentoComposeForm.vessel_name || undefined,
        subject: bentoComposeForm.subject_text,
        body: bentoComposeForm.body,
        recipient: FIXED_RECIPIENT,
        status: 'pending',
        attachments_count: bentoComposeForm.existing_attachment ? 1 : 0,
        attachment_names: bentoComposeForm.existing_attachment ? [bentoComposeForm.existing_attachment] : [],
        created_at: new Date().toISOString(),
      };
      host.setState({ bentoLogs: [pendingLog, ...bentoLogs] });

      try {
        const fd = new FormData();
        fd.append('vessel_name', bentoComposeForm.vessel_name);
        fd.append('datasource_tag', bentoComposeForm.datasource_tag);
        fd.append('subject_text', bentoComposeForm.subject_text);
        fd.append('body', bentoComposeForm.body);
        fd.append('recipient', FIXED_RECIPIENT);

       // Fetch the actual file bytes and attach as a real file
        if (bentoComposeForm.existing_attachment) {
          const attachFileId = host._getFileIdForAttachment(bentoComposeForm.vessel_name, bentoComposeForm.existing_attachment);

          if (!attachFileId) {
            host.setState(prev => ({
              bentoComposeBusy: false,
              bentoComposeErr: `Couldn't locate "${bentoComposeForm.existing_attachment}" as a real file — please reselect it from the list.`,
              bentoLogs: prev.bentoLogs.filter(l => l.id !== tempId),
            }));
            return;
          }

          fd.append('file_id', attachFileId);
          try {
            const fileRes = await fetch(`${host._base()}/api/files/${encodeURIComponent(attachFileId)}/content`, {
              headers: host._uploadHeaders(),
            });
            if (fileRes.ok) {
              const blob = await fileRes.blob();
              fd.append('attachment', blob, bentoComposeForm.existing_attachment);
            } else {
              fd.append('attachment_name', bentoComposeForm.existing_attachment);
            }
          } catch {
            fd.append('attachment_name', bentoComposeForm.existing_attachment);
          }
        }

        const res = await fetch(`${host._base()}/api/bento/dispatch`, {
          method: 'POST', headers: host._uploadHeaders(), body: fd,
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json().catch(() => ({}));

        // Update the pending log to completed
        host.setState(prev => ({
          bentoLogs: prev.bentoLogs.map(l =>
            l.id === tempId
              ? { ...l, id: data?.id || tempId, status: 'completed', sent_at: new Date().toISOString(), display_status: 'Sent' }
              : l
          ),
          bentoComposeBusy: false,
          bentoComposeMsg: `✅ Email dispatched to ${FIXED_RECIPIENT}`,
          bentoComposeForm: { vessel_name: '', datasource_tag: 'mail', subject_text: '[DataSource:mail]', body: '', file: null, existing_attachment: '', recipient: FIXED_RECIPIENT },
        }));
        setTimeout(() => host.setState({ bentoComposeOpen: false, bentoComposeMsg: null }), 2000);
      } catch (e: any) {
        // Mark log as failed
        host.setState(prev => ({
          bentoLogs: prev.bentoLogs.map(l =>
            l.id === tempId ? { ...l, status: 'failed', error_message: e?.message } : l
          ),
          bentoComposeBusy: false,
          bentoComposeErr: `Failed: ${e?.message || 'Unknown error'}`,
        }));
      }
    };

    return (
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, padding: isMobile ? 10 : 20 }}
        onClick={() => !bentoComposeBusy && host.setState({ bentoComposeOpen: false })}>
        <div style={{ background: '#fff', borderRadius: 12, padding: isMobile ? '16px 14px' : '28px 32px', width: isMobile ? '95vw' : 540, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,0.25)' }}
          onClick={e => e.stopPropagation()}>
          <div style={{ fontSize: 17, fontWeight: 800, color: '#0f172a', marginBottom: 20, display: 'flex', alignItems: 'center', gap: 8 }}>
            ✉ Compose & Dispatch Email
          </div>

          {bentoComposeMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{bentoComposeMsg}</div>}
          {bentoComposeErr && <div style={{ background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{bentoComposeErr}</div>}

          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>

            {/* Recipient – fixed, read-only */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Recipient (Fixed)</label>
              <input type="text" value={FIXED_RECIPIENT} readOnly
                style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#f8fafc', color: '#64748b', boxSizing: 'border-box', cursor: 'not-allowed' }} />
            </div>

            {/* Vessel Dropdown */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Vessel</label>
              <select value={bentoComposeForm.vessel_name}
                onChange={e => {
                  const selectedVesselName = e.target.value;
                  setForm({ vessel_name: selectedVesselName, existing_attachment: '' }, true);
                  if (selectedVesselName) {
                    // Fetch approved files directly from backend
                    host._fetchApprovedFilesForVessel(selectedVesselName).catch(() => undefined);
                    // Also load via existing mechanism as fallback
                    if (!host._filesLoadedForVessels.has(selectedVesselName)) {
                      host._filesLoadedForVessels.add(selectedVesselName);
                      host._loadFilesForVessel(selectedVesselName).catch(() => undefined);
                    }
                  }
                }}
                disabled={bentoComposeBusy}
                style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none' }}>
                <option value="">— Select vessel —</option>
                {vessels.map(v => <option key={v.id} value={v.name}>{v.name}</option>)}
              </select>
            </div>

            {/* Document Tag Dropdown */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Document Tag</label>
              <select value={bentoComposeForm.datasource_tag}
                onChange={e => setForm({ datasource_tag: e.target.value }, true)}
                disabled={bentoComposeBusy}
                style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none' }}>
                {Object.keys(DATASOURCE_TAGS_MAP).map(k => (
                  <option key={k} value={k}>{DATASOURCE_TAGS_MAP[k]} ({k})</option>
                ))}
              </select>
            </div>

            {/* Subject – auto-filled, editable [DataSource:TAG] Vessel Name / Subject */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>
                Subject * <span style={{ fontWeight: 400, color: '#94a3b8', fontSize: 11 }}>(auto-filled, editable)</span>
              </label>
              <input type="text"
                value={bentoComposeForm.subject_text}
                onChange={e => {
                  const val = e.target.value;
                  let suffix = val;
                  const slashIdx = val.indexOf(' / ');
                  if (slashIdx !== -1) {
                    suffix = val.substring(slashIdx + 3);
                  }
                  setForm({ subject_text: val, subject_suffix: suffix } as any);
                }}
                disabled={bentoComposeBusy}
                placeholder="[DataSource:TAG] Vessel Name / Subject..."
                style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #0284c7', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: '#f0f9ff' }} />
            </div>

            {/* Body */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Body</label>
              <textarea value={bentoComposeForm.body}
                onChange={e => setForm({ body: e.target.value })}
                disabled={bentoComposeBusy}
                placeholder="Email body..."
                rows={4}
                style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' }} />
            </div>

            {/* Attach Approved File Dropdown */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>
                Attach Approved File
                <span style={{ fontWeight: 400, color: '#94a3b8', fontSize: 11, marginLeft: 6 }}>(from selected vessel documents)</span>
              </label>
              <select
                value={bentoComposeForm.existing_attachment || ''}
                onChange={e => setForm({ existing_attachment: e.target.value })}
                disabled={bentoComposeBusy || !bentoComposeForm.vessel_name || isLoadingFiles}
                style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: bentoComposeForm.vessel_name ? '#fff' : '#f8fafc', outline: 'none' }}>
                {!bentoComposeForm.vessel_name ? (
                  <option value="">— Select Vessel First —</option>
                ) : isLoadingFiles ? (
                  <option value="">⏳ Loading files... —</option>
                ) : approvedFiles.length > 0 ? (
                  <>
                    <option value="">— Select Attached File —</option>
                    {approvedFiles.map(f => <option key={f} value={f}>📎 {f}</option>)}
                  </>
                ) : (
                  <option value="">— No approved files found for host vessel —</option>
                )}
              </select>
            </div>

            {/* Status indicator */}
            <div style={{ background: '#f8fafc', borderRadius: 6, padding: '8px 12px', fontSize: 11, color: '#64748b', border: '1px solid #e2e8f0' }}>
              📊 Status flow: <strong>Pending</strong> (on upload) → <strong>Pending</strong> (awaiting send) → <strong>Completed</strong> (after dispatch)
            </div>
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20, flexDirection: isMobile ? 'column' : 'row' }}>
            <button onClick={() => host.setState({ bentoComposeOpen: false, bentoComposeErr: null, bentoComposeMsg: null })}
              disabled={bentoComposeBusy}
              style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 16px', fontSize: 13, cursor: 'pointer' }}>Cancel</button>
            <button onClick={handleSend} disabled={bentoComposeBusy}
              style={{ background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              {bentoComposeBusy ? '⏳ Sending...' : '✉ Send Email'}
            </button>
          </div>
        </div>
      </div>
    );

}
