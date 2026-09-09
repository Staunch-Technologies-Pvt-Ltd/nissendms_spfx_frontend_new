import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { isMobileWidth } from '../responsive';

/** A single dialog supports both one-vessel and bulk temporary deletion. */
export function renderDeleteModal(host: VesselEmail): React.ReactElement {
  const { deleteVesselIds, deleteVesselProgress, vessels, modalBusy, modalMsg, modalError, deleteAutoCloseSeconds } = host.state;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
  const selected = vessels.filter(v => deleteVesselIds.has(v.id));
  const selectedCount = selected.length;
  const isSuccess = !!modalMsg && !modalBusy && !modalError;
  const progressEntries = Object.values(deleteVesselProgress || {});

  const getStatusVisual = (status?: 'waiting' | 'deleting' | 'success' | 'pending' | 'failed'): { icon: string; label: string; color: string } => {
    switch (status) {
      case 'success': return { icon: '✓', label: 'Deleted', color: '#15803d' };
      case 'deleting': return { icon: '…', label: 'Deleting', color: '#0284c7' };
      case 'pending': return { icon: '⏳', label: 'Pending approval', color: '#b45309' };
      case 'failed': return { icon: '✗', label: 'Failed', color: '#b91c1c' };
      default: return { icon: '○', label: 'Waiting', color: '#64748b' };
    }
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, padding: isMobile ? 10 : 20 }} onClick={host._closeModal}>
      <div style={{ background: '#fff', borderRadius: 10, padding: isMobile ? '16px' : '24px 28px', width: isMobile ? '95vw' : 520, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        {isSuccess ? (
          <div style={{ textAlign: 'center', padding: '16px 0 8px' }}>
            <div style={{ fontSize: 50, marginBottom: 10 }}>✅</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: '#059669', marginBottom: 8 }}>{selectedCount > 1 ? 'Vessels Deleted Successfully' : 'Vessel Deleted Successfully'}</div>
            <p style={{ fontSize: 13, color: '#475569', margin: '0 0 20px', lineHeight: 1.5 }}>{modalMsg}</p>
            {progressEntries.length > 0 && (
              <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 12px', margin: '0 0 16px', textAlign: 'left', maxHeight: 180, overflowY: 'auto' }}>
                {progressEntries.filter(p => p.status === 'success').map((p, idx) => (
                  <div key={`${p.name}_${idx}`} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, padding: '6px 0', borderBottom: idx === progressEntries.filter(x => x.status === 'success').length - 1 ? 'none' : '1px solid #e2e8f0' }}>
                    <span style={{ fontSize: 13, color: '#0f172a', fontWeight: 600 }}>{p.name}</span>
                    <span style={{ fontSize: 12, color: '#15803d', fontWeight: 700 }}>✓ Deleted</span>
                  </div>
                ))}
              </div>
            )}
            {typeof deleteAutoCloseSeconds === 'number' && (
              <div style={{ margin: '0 0 14px', fontSize: 12, color: '#0f766e', fontWeight: 600 }}>
                Auto-closing in {deleteAutoCloseSeconds}s...
              </div>
            )}
            <button
              onClick={host._closeModal}
              style={{ background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 8, minHeight: 44, padding: '10px 28px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}
            >
              Close
            </button>
          </div>
        ) : (
          <>
        <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 8 }}>Delete vessel{selectedCount > 1 ? 's' : ''}</div>
        <p style={{ fontSize: 13, color: '#475569', margin: '0 0 14px' }}>
          This is temporary: the vessel folder and its contents will move to the Recycle Bin and can be restored.
        </p>
        {selectedCount > 0 ? (
          <div style={{ background: '#fff1f2', border: '1px solid #fecdd3', borderRadius: 8, padding: '10px 12px', marginBottom: 14 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#9f1239', marginBottom: 8 }}>{selectedCount} vessel{selectedCount === 1 ? '' : 's'} selected for deletion</div>
            <div style={{ maxHeight: 220, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 8 }}>
              {selected.map(v => (
                <div
                  key={v.id}
                  style={{
                    background: deleteVesselProgress?.[v.id]?.status === 'success' ? '#ecfdf5' : '#fff',
                    border: deleteVesselProgress?.[v.id]?.status === 'success' ? '1px solid #86efac' : '1px solid #fecdd3',
                    borderRadius: 8,
                    padding: '8px 10px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 12,
                  }}
                >
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: '#881337' }}>{v.name}</div>
                    <div style={{ marginTop: 2, fontSize: 11, color: '#9f1239' }}>{v.vessel_type || 'Vessel'}{v.imo ? ` | IMO ${v.imo}` : ''}</div>
                  </div>
                  <div style={{ textAlign: 'right', minWidth: 110 }}>
                    {(() => {
                      const info = getStatusVisual(deleteVesselProgress?.[v.id]?.status);
                      return <span style={{ fontSize: 12, color: info.color, fontWeight: 700 }}>{info.icon} {info.label}</span>;
                    })()}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div style={{ padding: 16, color: '#64748b', fontSize: 13 }}>No vessel selected.</div>
        )}
        {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalMsg}</div>}
        {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalError}</div>}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, flexDirection: isMobile ? 'column' : 'row' }}>
          <button style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 16px', fontSize: 13, cursor: 'pointer' }} onClick={host._closeModal} disabled={modalBusy}>Cancel</button>
          <button style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: selectedCount > 0 ? 'pointer' : 'not-allowed', opacity: selectedCount > 0 ? 1 : 0.55 }} onClick={host._submitDelete} disabled={modalBusy || selectedCount === 0}>
            {modalBusy ? 'Moving to Recycle Bin...' : 'Delete and Move to Recycle Bin'}
          </button>
        </div>
          </>
        )}
      </div>
    </div>
  );
}
