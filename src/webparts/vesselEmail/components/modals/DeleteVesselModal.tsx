import * as React from 'react';
import type VesselEmail from '../VesselEmail';

/** A single dialog supports both one-vessel and bulk temporary deletion. */
export function renderDeleteModal(host: VesselEmail): React.ReactElement {
  const { selectedVessel, modalBusy, modalMsg, modalError } = host.state;
  const isSuccess = !!modalMsg && !modalBusy && !modalError;

  React.useEffect(() => {
    if (!isSuccess) return undefined;
    const timer = setTimeout(() => host._closeModal(), 15000);
    return () => clearTimeout(timer);
  }, [isSuccess, host]);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={host._closeModal}>
      <div style={{ background: '#fff', borderRadius: 10, padding: '24px 28px', width: 520, maxWidth: '90vw', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        {isSuccess ? (
          <div style={{ textAlign: 'center', padding: '16px 0 8px' }}>
            <div style={{ fontSize: 50, marginBottom: 10 }}>✅</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: '#059669', marginBottom: 8 }}>Vessel Deleted Successfully</div>
            <p style={{ fontSize: 13, color: '#475569', margin: '0 0 20px', lineHeight: 1.5 }}>{modalMsg}</p>
            <button
              onClick={host._closeModal}
              style={{ background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 28px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}
            >
              Done
            </button>
          </div>
        ) : (
          <>
        <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 8 }}>Delete vessel</div>
        <p style={{ fontSize: 13, color: '#475569', margin: '0 0 14px' }}>
          This is temporary: the vessel folder and its contents will move to the Recycle Bin and can be restored.
        </p>
        {selectedVessel ? (
          <div style={{ background: '#fff1f2', border: '1px solid #fecdd3', borderRadius: 8, padding: '14px 16px', marginBottom: 14 }}>
            <div style={{ fontSize: 15, fontWeight: 700, color: '#881337' }}>{selectedVessel.name}</div>
            <div style={{ marginTop: 4, fontSize: 11, color: '#9f1239' }}>{selectedVessel.vessel_type || 'Vessel'}{selectedVessel.imo ? ` | IMO ${selectedVessel.imo}` : ''}</div>
          </div>
        ) : (
          <div style={{ padding: 16, color: '#64748b', fontSize: 13 }}>No vessel selected.</div>
        )}
        {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalMsg}</div>}
        {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalError}</div>}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }} onClick={host._closeModal} disabled={modalBusy}>Cancel</button>
          <button style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: selectedVessel ? 'pointer' : 'not-allowed', opacity: selectedVessel ? 1 : 0.55 }} onClick={host._submitDelete} disabled={modalBusy || !selectedVessel}>
            {modalBusy ? 'Moving to Recycle Bin...' : 'Delete and Move to Recycle Bin'}
          </button>
        </div>
          </>
        )}
      </div>
    </div>
  );
}
