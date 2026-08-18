import * as React from 'react';
import type VesselEmail from '../VesselEmail';

/** A single dialog supports both one-vessel and bulk temporary deletion. */
export function renderDeleteModal(host: VesselEmail): React.ReactElement {
  const { vessels, deleteVesselIds, modalBusy, modalMsg, modalError } = host.state;
  const selectedCount = deleteVesselIds.size;
  const allSelected = vessels.length > 0 && vessels.every(v => deleteVesselIds.has(v.id));

  const toggleVessel = (id: string): void => {
    host.setState(prev => {
      const next = new Set(prev.deleteVesselIds);
      if (next.has(id)) next.delete(id); else next.add(id);
      return { deleteVesselIds: next, modalError: null };
    });
  };

  const toggleAll = (): void => {
    host.setState({
      deleteVesselIds: allSelected ? new Set() : new Set(vessels.map(v => v.id)),
      modalError: null,
    });
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={host._closeModal}>
      <div style={{ background: '#fff', borderRadius: 10, padding: '24px 28px', width: 640, maxWidth: '90vw', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 8 }}>Delete vessels</div>
        <p style={{ fontSize: 13, color: '#475569', margin: '0 0 14px' }}>
          Select one or more vessels. This is temporary: the vessel folder and its contents move to the Recycle Bin and can be restored.
        </p>
        <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden', marginBottom: 14 }}>
          <label style={{ display: 'flex', gap: 10, alignItems: 'center', padding: '10px 12px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 13, fontWeight: 700, cursor: modalBusy ? 'default' : 'pointer' }}>
            <input type="checkbox" checked={allSelected} onChange={toggleAll} disabled={modalBusy} />
            Select all created vessels ({vessels.length})
          </label>
          <div style={{ maxHeight: 260, overflowY: 'auto' }}>
            {vessels.length === 0 ? (
              <div style={{ padding: 16, color: '#64748b', fontSize: 13 }}>No vessels are available to delete.</div>
            ) : vessels.map(vessel => (
              <label key={vessel.id} style={{ display: 'flex', gap: 10, alignItems: 'center', padding: '10px 12px', borderBottom: '1px solid #f1f5f9', fontSize: 13, cursor: modalBusy ? 'default' : 'pointer', background: deleteVesselIds.has(vessel.id) ? '#fff1f2' : '#fff' }}>
                <input type="checkbox" checked={deleteVesselIds.has(vessel.id)} onChange={() => toggleVessel(vessel.id)} disabled={modalBusy} />
                <span style={{ flex: 1, fontWeight: 600, color: '#0f172a' }}>{vessel.name}</span>
                <span style={{ fontSize: 11, color: '#64748b' }}>{vessel.vessel_type || 'Vessel'}{vessel.imo ? ` | IMO ${vessel.imo}` : ''}</span>
              </label>
            ))}
          </div>
        </div>
        {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalMsg}</div>}
        {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalError}</div>}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }} onClick={host._closeModal} disabled={modalBusy}>Cancel</button>
          <button style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: selectedCount ? 'pointer' : 'not-allowed', opacity: selectedCount ? 1 : 0.55 }} onClick={host._submitDelete} disabled={modalBusy || selectedCount === 0}>
            {modalBusy ? 'Moving to Recycle Bin...' : `Delete selected (${selectedCount})`}
          </button>
        </div>
      </div>
    </div>
  );
}
