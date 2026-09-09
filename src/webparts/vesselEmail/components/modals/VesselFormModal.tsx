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

export function renderVesselForm(host: VesselEmail, mode: 'create' | 'edit'): React.ReactElement {
  return <VesselFormContent host={host} mode={mode} />;
}

function VesselFormContent({ host, mode }: { host: VesselEmail; mode: 'create' | 'edit' }): React.ReactElement {
  const { form, modalBusy, modalMsg, modalError, formFieldErrors, vessels } = host.state;  const isCreate = mode === 'create';
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
  const [elapsed, setElapsed] = React.useState(0);

  React.useEffect(() => {
    if (!modalBusy) {
      setElapsed(0);
      return;
    }
    const timer = setInterval(() => setElapsed(e => e + 1), 1000);
    return () => clearInterval(timer);
  }, [modalBusy]);

  const formatTimer = (sec: number): string => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${m < 10 ? '0' + m : m}:${s < 10 ? '0' + s : s}`;
  };

  const set = (k: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
    host.setState({ form: { ...form, [k]: e.target.value }, modalError: null, formFieldErrors: { ...formFieldErrors, [k]: '' } });

  const setName = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const value = e.target.value;
    const normalized = value.replace(/[ _'\"]+/g, '').toLowerCase();
    const duplicate = normalized.length > 0 && vessels.some(v =>
      v.id !== host.state.selectedVessel?.id &&
      v.name.replace(/[ _'\"]+/g, '').toLowerCase() === normalized
    );
    host.setState({
      form: { ...form, name: value },
      modalError: null,
      formFieldErrors: { ...formFieldErrors, name: duplicate ? 'Vessel name already exists.' : '' },
    });
  };

  const isSuccess = modalMsg && modalMsg.startsWith('🎉');
  const isEditSuccess = !isCreate && !!modalMsg && (modalMsg.startsWith('✅') || modalMsg.startsWith('⏳'));

   const handleClose = () => {
    host.setState({ modal: 'none', modalMsg: null, modalError: null, formFieldErrors: {} });
    if (isSuccess || isEditSuccess) {
      host._loadData();
    }
  };

  return (
    <div
      style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, padding: isMobile ? 10 : 20 }}
      onClick={e => { if (e.target === e.currentTarget && !modalBusy) handleClose(); }}
    >
      <div style={{ background: '#fff', borderRadius: 20, padding: isMobile ? '20px 16px' : '32px 36px', width: isMobile ? '95vw' : 480, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 24px 64px rgba(0,0,0,0.28)', position: 'relative' }}>

        {!modalBusy && (
          <button
            onClick={handleClose}
            style={{ position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, color: '#94a3b8', cursor: 'pointer' }}
            title="Close"
          >✕</button>
        )}

        {/* ── STATE 1: PROVISIONING IN PROGRESS (Timer & Progress Bar) ── */}
        {modalBusy && !isCreate ? (
          <div style={{ textAlign: 'center', padding: '28px 0' }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0284c7' }}>
              Updating Vessel
            </h3>
            <p style={{ margin: 0, fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              Saving the updated vessel details. SharePoint folders will not be provisioned.
            </p>
          </div>
        ) : modalBusy ? (
          <div style={{ textAlign: 'center', padding: '12px 0' }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div>
            <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' }}>
              Creating & Provisioning Vessel…
            </h3>
            <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              Registering <strong>"{form.name || 'Vessel'}"</strong> in DMS database & provisioning SharePoint folder structure. Please wait.
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
                height: '100%', width: `${Math.min(96, 20 + elapsed * 15)}%`,
                transition: 'width 0.8s ease-out', borderRadius: 10,
              }} />
            </div>

            {/* Step Progress Checklist */}
            <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 }}>
              <div style={{ color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 14 }}>✓</span><span>Vessel record registered in database</span>
              </div>
              <div style={{ color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 14 }}>⏳</span><span>Creating SharePoint DMS folder tree (Technical & Crewing, Commercial, Insurance)…</span>
              </div>
              <div style={{ color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 14 }}>○</span><span>Linking category subfolders & permissions</span>
              </div>
            </div>
          </div>
        ) : isSuccess ? (
          /* ── STATE 2: SUCCESS SCREEN ── */
          <div style={{ textAlign: 'center', padding: '12px 0' }}>
            <div style={{ fontSize: 52, marginBottom: 12 }}>🎉</div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' }}>
              Vessel Successfully Created & Provisioned!
            </h3>
            <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              <strong>"{form.name}"</strong> has been registered in the DMS database and its full SharePoint folder structure has been created.
            </p>

            <div style={{ background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '14px 16px', marginBottom: 24, fontSize: 12, color: '#065f46', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, marginBottom: 4 }}>✅ Vessel Ready</div>
              <div>• Registered vessel card added to main grid</div>
              <div>• IMO Number: {form.imo || '—'}</div>
              <div>• SharePoint DMS department subfolders created</div>
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
        ) : isEditSuccess ? (
          <div style={{ textAlign: 'center', padding: '28px 0' }}>
            <div style={{ fontSize: 52, marginBottom: 12 }}>✅</div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' }}>
              Vessel Updated Successfully
            </h3>
            <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              The vessel details were saved without changing its SharePoint folders.
            </p>
            <button
              onClick={handleClose}
              style={{
                background: 'linear-gradient(135deg, #10b981, #059669)',
                color: '#fff', border: 'none', borderRadius: 10,
                padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
                boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
              }}
            >
              Done
            </button>
          </div>
        ) : (
          /* ── STATE 3: FORM ENTRY ── */
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
                <div style={{ width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>
                  🚢
                </div>
                <div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: '#0f172a' }}>{isCreate ? 'New Vessel' : 'Update Vessel'}</div>
                  <div style={{ fontSize: 12, color: '#64748b', marginTop: 2, maxWidth: 280 }}>
                    {isCreate ? 'Provisions the full folder structure across all 3 main folders.' : 'Update the vessel details below.'}
                  </div>
                </div>
              </div>
            </div>

            {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalMsg}</div>}
            {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalError}</div>}

            <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
              <div>
                <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>
                  Vessel name <span style={{ color: '#ef4444' }}>*</span>
                </label>
                <input
                  style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: `1px solid ${formFieldErrors?.name ? '#ef4444' : '#cbd5e1'}`, fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                  value={form.name} onChange={setName} placeholder="e.g. MV Pacific Trader" />
                {formFieldErrors?.name && (
                  <div style={{ marginTop: 6, fontSize: 12, color: '#dc2626', display: 'flex', alignItems: 'center', gap: 4 }}>
                    <span>⚠️</span><span>{formFieldErrors.name}</span>
                  </div>
                )}
              </div>

              <div>
                <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>
                  IMO number <span style={{ color: '#ef4444' }}>*</span>
                </label>
                             <input
                  style={{
                    width: '100%', padding: '10px 14px', borderRadius: 8,
                    border: `1px solid ${formFieldErrors?.imo ? '#ef4444' : '#cbd5e1'}`,
                    fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a',
                  }}
                  value={form.imo}
                  onChange={e => {
                    const val = e.target.value.replace(/\D/g, '').slice(0, 7);
                    const duplicate = val.length > 0 && vessels.some(v =>
                      v.id !== host.state.selectedVessel?.id && (v.imo || '').trim() === val
                    );
                    host.setState({
                      form: { ...form, imo: val },
                      modalError: null,
                      formFieldErrors: { ...formFieldErrors, imo: duplicate ? 'A vessel with that IMO number already exists.' : '' },
                    });
                  }}
                  placeholder="7 digits, e.g. 9074729"
                  maxLength={7}
                  inputMode="numeric" />
                {formFieldErrors?.imo && (
                  <div style={{ marginTop: 6, fontSize: 12, color: '#dc2626', display: 'flex', alignItems: 'center', gap: 4 }}>
                    <span>⚠️</span><span>{formFieldErrors.imo}</span>
                  </div>
                )}
              </div>

              <div>
                <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Ship yard name</label>
                <input
                  style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                  value={form.shipyard} onChange={set('shipyard')} placeholder="e.g. Hyundai Heavy Industries" />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Hull number</label>
                <input
                  style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                  value={form.hull_number} onChange={set('hull_number')} placeholder="e.g. H2456" />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Vessel type</label>
                <select
                  style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', background: '#fff', color: form.vessel_type ? '#0f172a' : '#94a3b8', appearance: 'auto' }}
                  value={form.vessel_type} onChange={set('vessel_type')}>
                  <option value="">Select a type...</option>
                  {VESSEL_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', flexDirection: isMobile ? 'column' : 'row', gap: 10, marginTop: 28 }}>
              <button
                style={{ background: 'transparent', border: '1px solid #cbd5e1', borderRadius: 8, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '10px 20px', fontSize: 14, fontWeight: 500, color: '#475569', cursor: 'pointer' }}
                onClick={handleClose}>Cancel</button>
              <button
                style={{ background: 'linear-gradient(135deg, #0d9488, #0f766e)', color: '#fff', border: 'none', borderRadius: 8, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}
                onClick={isCreate ? host._submitCreate : host._submitEdit}>
                {isCreate ? 'Create Vessel' : 'Update Vessel'}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
