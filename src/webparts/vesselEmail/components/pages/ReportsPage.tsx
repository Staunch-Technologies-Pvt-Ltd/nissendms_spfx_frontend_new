import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge, cleanName } from '../constants';
import { extractRealDocuments, RealDashboardDoc } from './DashboardPage';

export function renderReportsPage(host: VesselEmail): React.ReactElement {
  const { vessels, reportsSelectedVessel } = host.state;
  const selectedVessel = reportsSelectedVessel || 'All Vessels';

    const allRealDocs = extractRealDocuments(host);
    const filteredDocs = selectedVessel === 'All Vessels'
      ? allRealDocs
      : allRealDocs.filter(d => (d.vessel || '').toLowerCase() === selectedVessel.toLowerCase());

    const totalDocs = filteredDocs.length;
    const expiredDocs = filteredDocs.filter(d => d.status === 'Expired').length;
    const expiringSoonDocs = filteredDocs.filter(d => d.status === 'Expiring Soon').length;
    const validDocs = filteredDocs.filter(d => d.status === 'Valid').length;

    // Distribution by Type
    const typeCounts = new Map<string, number>();
    filteredDocs.forEach(d => {
      const t = d.type || 'Other';
      typeCounts.set(t, (typeCounts.get(t) || 0) + 1);
    });
    const typeList = Array.from(typeCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);

    // Distribution by Vessel
    const vesselCounts = new Map<string, number>();
    (vessels || []).forEach(v => vesselCounts.set(v.name, 0));
    allRealDocs.forEach(d => {
      if (d.vessel && d.vessel !== 'Shared Documents') {
        vesselCounts.set(d.vessel, (vesselCounts.get(d.vessel) || 0) + 1);
      }
    });
    const vesselList = Array.from(vesselCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    const maxVesselDocs = Math.max(1, ...vesselList.map(v => v[1]));

    const typeColors = ['#2563eb', '#f59e0b', '#10b981', '#8b5cf6', '#06b6d4'];

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Reports & Compliance</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Live compliance summary and document analytics across your fleet.</p>
          </div>

          <button
            onClick={() => {
              window.print();
            }}
            style={{
              background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6,
              padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              display: 'inline-flex', alignItems: 'center', gap: 6,
              boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
            }}
          >
            🖨️ Print / Export Summary
          </button>
        </div>

        {/* Date / Filter bar */}
        <div style={{ background: '#fff', borderRadius: 10, padding: 14, border: '1px solid #e2e8f0', display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: '#334155' }}>
            <span>📅</span>
            <span style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>Live Fleet Data</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 'auto' }}>
            <span style={{ fontSize: 12, color: '#64748b', fontWeight: 600 }}>Filter by Vessel:</span>
            <select
              value={selectedVessel}
              onChange={e => host.setState({ reportsSelectedVessel: e.target.value })}
              style={{ border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none', background: '#fff', fontWeight: 500 }}
            >
              <option value="All Vessels">All Vessels ({allRealDocs.length} docs)</option>
              {(vessels || []).map(v => (
                <option key={v.id} value={v.name}>{v.name}</option>
              ))}
            </select>
          </div>
        </div>

        {/* 4 Stat Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16 }}>
          <div style={{ background: '#fff', borderRadius: 10, padding: 18, border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Total Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#0284c7', marginTop: 4 }}>{totalDocs}</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 10, padding: 18, border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Expired Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: expiredDocs > 0 ? '#ef4444' : '#0f172a', marginTop: 4 }}>{expiredDocs}</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 10, padding: 18, border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Expiring in 30 days</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: expiringSoonDocs > 0 ? '#f59e0b' : '#0f172a', marginTop: 4 }}>{expiringSoonDocs}</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 10, padding: 18, border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Valid Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#10b981', marginTop: 4 }}>{validDocs}</div>
          </div>
        </div>

        {/* Charts Grid */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          {/* Documents by Type */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20, boxShadow: '0 1px 4px rgba(0,0,0,0.03)' }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Documents by Category / Type</h3>
            {typeList.length === 0 ? (
              <div style={{ padding: '32px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                No document category data available yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10, fontSize: 13 }}>
                {typeList.map(([tName, tCount], idx) => {
                  const pct = totalDocs > 0 ? Math.round((tCount / totalDocs) * 100) : 0;
                  const color = typeColors[idx % typeColors.length];
                  return (
                    <div key={tName}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                        <span style={{ color: '#334155', fontWeight: 600 }}>{tName}</span>
                        <span style={{ fontWeight: 700, color: '#0f172a' }}>{tCount} ({pct}%)</span>
                      </div>
                      <div style={{ background: '#f1f5f9', borderRadius: 4, height: 10, overflow: 'hidden' }}>
                        <div style={{ width: `${pct}%`, background: color, height: '100%', borderRadius: 4, transition: 'width 0.3s' }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Documents by Vessel (Bar Chart) */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20, boxShadow: '0 1px 4px rgba(0,0,0,0.03)' }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Documents by Vessel</h3>
            {vesselList.length === 0 ? (
              <div style={{ padding: '32px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                No vessel document data available yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {vesselList.map(([vName, vCount]) => (
                  <div key={vName} style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: 12 }}>
                    <span style={{ width: 110, color: '#475569', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={vName}>
                      🚢 {vName}
                    </span>
                    <div style={{ flex: 1, background: '#f1f5f9', borderRadius: 4, height: 16, overflow: 'hidden' }}>
                      <div style={{ width: `${(vCount / maxVesselDocs) * 100}%`, background: '#0284c7', height: '100%', borderRadius: 4, transition: 'width 0.3s' }} />
                    </div>
                    <span style={{ width: 28, fontWeight: 700, textAlign: 'right', color: '#0f172a' }}>{vCount}</span>
                  </div>
                ))}
                <div style={{ fontSize: 10, color: '#94a3b8', textAlign: 'center', marginTop: 8 }}>No. of Documents</div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
}
