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

function getSpoVesselFolderUrl(siteUrlProp?: string, vesselName?: string): string {
  if (!siteUrlProp) return '#';
  try {
    const urlObj = new URL(siteUrlProp);
    const basePath = urlObj.pathname.replace(/\/$/, '');
    const folderPath = `${basePath}/Shared Documents/Vessel Management/Technical & Crewing${vesselName ? '/' + vesselName : ''}`;
    return `${urlObj.origin}${basePath}/Shared Documents/Forms/AllItems.aspx?id=${encodeURIComponent(folderPath)}`;
  } catch {
    return '#';
  }
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

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>

        {/* ── Header ── */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20, flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
              🚢 Vessels
              {vessels.length > 0 && (
                <span style={{ background: '#e0f2fe', color: '#0284c7', borderRadius: 20, padding: '2px 10px', fontSize: 13, fontWeight: 700 }}>
                  {vessels.length}
                </span>
              )}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              Manage fleet vessels, provision SharePoint folders, and view documents.
            </p>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button
              onClick={() => host._goToView('vessels').catch(() => undefined)}
              title="Reload vessel list from database"
              style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '9px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}
            >
              🔄 Refresh
            </button>
            <button
              onClick={() => selectedVessel ? host._openDeleteVessel(selectedVessel) : alert('Please select a vessel first.')}
              style={{ background: 'linear-gradient(135deg, #f43f5e, #e11d48)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}
            >
              🗑 Delete
            </button>
            <button
              onClick={() => selectedVessel ? host._openEditVessel(selectedVessel) : alert('Please select a vessel first.')}
              style={{ background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}
            >
              ✏️ Edit
            </button>
            <button
              onClick={host._openCreate}
              style={{ background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}
            >
              ＋ New Vessel
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
                ＋ Create First Vessel
              </button>
            )}
          </div>
        ) : (
          /* ── Vessel Cards Grid ── */
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: 16 }}>
            {filtered.map(vessel => {
              const isSelected = selectedVessel?.id === vessel.id;
              const status = vessel.status || 'Active';
              const statusColor = status === 'Active' ? '#10b981' : status === 'In Maintenance' ? '#f59e0b' : '#ef4444';
              const statusBg = status === 'Active' ? '#f0fdf4' : status === 'In Maintenance' ? '#fffbeb' : '#fef2f2';
              const isProvisioning = folderProvisioningVesselId === vessel.id;
              // Use deterministic vessel image from pool
              const imgSrc = getVesselImageForId(vessel.id);

              return (
                <div
                  key={vessel.id}
                  onClick={() => host.setState({ selectedVessel: isSelected ? null : vessel })}
                  style={{
                    background: '#fff',
                    borderRadius: 14,
                    border: isSelected ? '2px solid #0078d4' : '1px solid #e2e8f0',
                    boxShadow: isSelected ? '0 0 0 3px rgba(0,120,212,0.15)' : '0 1px 4px rgba(0,0,0,0.06)',
                    overflow: 'hidden',
                    cursor: 'pointer',
                    transition: 'all 0.18s ease',
                    display: 'flex',
                    flexDirection: 'column',
                  }}
                  onMouseEnter={e => { if (!isSelected) e.currentTarget.style.boxShadow = '0 4px 16px rgba(0,0,0,0.10)'; }}
                  onMouseLeave={e => { if (!isSelected) e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.06)'; }}
                >
                  {/* Card image header */}
                  <div style={{ position: 'relative', height: 120, overflow: 'hidden', background: '#1e3a5f' }}>
                    <img
                      src={resolveImgUrl(imgSrc)}
                      alt={vessel.name}
                      style={{ width: '100%', height: '100%', objectFit: 'cover', opacity: 0.85 }}
                    />
                    <div style={{
                      position: 'absolute', inset: 0,
                      background: 'linear-gradient(to bottom, transparent 30%, rgba(15,23,42,0.75) 100%)',
                    }} />
                    {/* Status badge */}
                    <span style={{
                      position: 'absolute', top: 10, right: 10,
                      background: statusBg, color: statusColor,
                      borderRadius: 20, padding: '2px 10px', fontSize: 11, fontWeight: 700,
                      border: `1px solid ${statusColor}40`,
                    }}>
                      {status}
                    </span>
                    {/* Selection indicator */}
                    {isSelected && (
                      <span style={{
                        position: 'absolute', top: 10, left: 10,
                        background: '#0078d4', color: '#fff',
                        borderRadius: '50%', width: 22, height: 22,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        fontSize: 13, fontWeight: 700, boxShadow: '0 2px 6px rgba(0,0,0,0.25)',
                      }}>
                        ✓
                      </span>
                    )}
                    {/* Vessel name over image */}
                    <div style={{ position: 'absolute', bottom: 10, left: 14, right: 14 }}>
                      <p style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#fff', textShadow: '0 1px 3px rgba(0,0,0,0.5)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                        🚢 {vessel.name}
                      </p>
                    </div>
                  </div>

                  {/* Card body */}
                  <div style={{ padding: '14px 16px', flex: 1, display: 'flex', flexDirection: 'column', gap: 8 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 12px' }}>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>IMO</span>
                        <p style={{ margin: 0, fontSize: 13, color: '#1e293b', fontWeight: 600 }}>{vessel.imo || '—'}</p>
                      </div>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>Type</span>
                        <p style={{ margin: 0, fontSize: 13, color: '#1e293b', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{vessel.vessel_type || '—'}</p>
                      </div>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>Shipyard</span>
                        <p style={{ margin: 0, fontSize: 12, color: '#475569', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{vessel.shipyard || '—'}</p>
                      </div>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>Hull No.</span>
                        <p style={{ margin: 0, fontSize: 12, color: '#475569' }}>{vessel.hull_number || '—'}</p>
                      </div>
                    </div>

                    {/* Action buttons */}
                    <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                      {/* View Documents — navigate to Documents filtered for host vessel */}
                      <button
                        onClick={e => {
                          e.stopPropagation();
                          host.setState({ vesselFilter: vessel.name, docMainFolder: null, folderPathStack: [] });
                          void host._goToView('list');
                        }}
                        title="View documents for host vessel"
                        style={{
                          flex: 1, background: '#eff6ff', color: '#1d4ed8',
                          border: '1px solid #bfdbfe', borderRadius: 7, padding: '7px 10px',
                          fontSize: 11, fontWeight: 700, cursor: 'pointer',
                          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                        }}
                      >
                        📄 View Documents
                      </button>
                      {/* Provision SPO folders */}
                      <button
                        disabled={!!folderProvisioningVesselId}
                        onClick={e => {
                          e.stopPropagation();
                          host._provisionVesselFolders(vessel.name, vessel.id).catch(() => undefined);
                        }}
                        title="Create SharePoint folder structure for host vessel"
                        style={{
                          flex: 1,
                          border: '1px solid #cbd5e1', borderRadius: 7, padding: '7px 10px',
                          fontSize: 11, fontWeight: 700, cursor: folderProvisioningVesselId ? 'not-allowed' : 'pointer',
                          background: isProvisioning ? '#f0f9ff' : '#f8fafc',
                          color: isProvisioning ? '#0284c7' : '#334155',
                          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                        }}
                      >
                        {isProvisioning ? (
                          <><span style={{ display: 'inline-block', width: 11, height: 11, border: '2px solid #0284c7', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} /> Creating…</>
                        ) : '📁 Provision'}
                      </button>
                      {/* Open exact vessel folder */}
                      <a
                        href={getSpoVesselFolderUrl(host.props.siteUrl, vessel.name)}
                        target="_blank"
                        rel="noopener noreferrer"
                        onClick={e => e.stopPropagation()}
                        title="Open in SharePoint"
                        style={{
                          width: 32, display: 'flex', alignItems: 'center', justifyContent: 'center',
                          border: '1px solid #cbd5e1', borderRadius: 7, background: '#f8fafc',
                          color: '#0078d4', fontSize: 13, textDecoration: 'none',
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

        {/* Keyframe animation for spinners */}
        <style>{`
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        `}</style>

        {/* Modals */}
        {modal === 'create' && host._renderVesselForm('create')}
        {modal === 'edit' && host._renderVesselForm('edit')}
        {modal === 'delete' && host._renderDeleteModal()}
      </div>
    );
}
