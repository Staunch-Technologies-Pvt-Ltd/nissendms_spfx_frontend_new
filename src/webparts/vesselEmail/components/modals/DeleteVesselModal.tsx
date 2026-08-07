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

export function renderDeleteModal(host: VesselEmail): React.ReactElement {
    const { selectedVessel, modalBusy, modalMsg, modalError } = host.state;
    return (
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={host._closeModal}>
        <div style={{ background: '#fff', borderRadius: 10, padding: '24px 28px', width: 400, maxWidth: '90vw', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
          <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 12 }}>🗑 Delete Vessel</div>
          <p style={{ fontSize: 13, color: '#475569', marginBottom: 16 }}>
            Are you sure you want to delete <strong>"{selectedVessel?.name}"</strong>?
          </p>
          {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalMsg}</div>}
          {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalError}</div>}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
            <button style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }} onClick={host._closeModal} disabled={modalBusy}>Cancel</button>
            <button style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }} onClick={host._submitDelete} disabled={modalBusy}>
              {modalBusy ? 'Deleting...' : 'Delete'}
            </button>
          </div>
        </div>
      </div>
    );
}
