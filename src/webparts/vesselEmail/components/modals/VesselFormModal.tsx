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

export function renderVesselForm(host: VesselEmail, mode: 'create' | 'edit'): React.ReactElement {
    const { form, modalBusy, modalMsg, modalError } = host.state;
    const set = (k: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
      host.setState({ form: { ...form, [k]: e.target.value }, modalError: null });
    const isCreate = mode === 'create';

    return (
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={host._closeModal}>
        <div style={{ background: '#fff', borderRadius: 16, padding: '32px 36px', width: 480, maxWidth: '92vw', boxShadow: '0 12px 40px rgba(0,0,0,0.18)' }} onClick={e => e.stopPropagation()}>

          {/* Modal Header */}
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
            <button onClick={host._closeModal} disabled={modalBusy}
              style={{ background: 'none', border: 'none', fontSize: 18, color: '#94a3b8', cursor: 'pointer', lineHeight: 1, padding: 4 }}>✕</button>
          </div>

          {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalMsg}</div>}
          {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalError}</div>}

          {/* Fields */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>
                Vessel name <span style={{ color: '#ef4444' }}>*</span>
              </label>
              <input
                style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                value={form.name} onChange={set('name')} placeholder="e.g. MV Pacific Trader" disabled={modalBusy} />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>
                IMO number <span style={{ color: '#ef4444' }}>*</span>
              </label>
              <input
                style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                value={form.imo}
                onChange={e => {
                  const val = e.target.value.replace(/\D/g, '').slice(0, 7);
                  host.setState({ form: { ...form, imo: val }, modalError: null });
                }}
                placeholder="7 digits, e.g. 9074729"
                maxLength={7}
                inputMode="numeric"
                disabled={modalBusy} />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Ship yard name</label>
              <input
                style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                value={form.shipyard} onChange={set('shipyard')} placeholder="e.g. Hyundai Heavy Industries" disabled={modalBusy} />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Hull number</label>
              <input
                style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }}
                value={form.hull_number} onChange={set('hull_number')} placeholder="e.g. H2456" disabled={modalBusy} />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Vessel type</label>
              <select
                style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', background: '#fff', color: form.vessel_type ? '#0f172a' : '#94a3b8', appearance: 'auto' }}
                value={form.vessel_type} onChange={set('vessel_type')} disabled={modalBusy}>
                <option value="">Select a type...</option>
                {VESSEL_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
          </div>

          {/* Footer */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 28 }}>
            <button
              style={{ background: 'transparent', border: 'none', borderRadius: 8, padding: '10px 20px', fontSize: 14, fontWeight: 500, color: '#475569', cursor: 'pointer' }}
              onClick={host._closeModal} disabled={modalBusy}>Cancel</button>
            <button
              style={{ background: 'linear-gradient(135deg, #0d9488, #0f766e)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}
              onClick={isCreate ? host._submitCreate : host._submitEdit} disabled={modalBusy}>
              {modalBusy ? 'Saving...' : isCreate ? 'Create Vessel' : 'Update Vessel'}
            </button>
          </div>
        </div>
      </div>
    );
}
