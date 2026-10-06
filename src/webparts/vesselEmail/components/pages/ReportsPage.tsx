import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge, cleanName } from '../constants';
import { extractRealDocuments, RealDashboardDoc } from './DashboardPage';
import { Icon } from '@fluentui/react/lib/Icon';
import { clay } from '../clayTheme';
import { DmsPageHeader, DMS_SECTION_CARD, dmsBtn, dmsTone } from '../dmsDesignSystem';

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

    // Category colors for the chart legends below — drawn from `clay.*` so the
    // palette stays token-backed and re-themes with the rest of the app.
    const typeColors = [clay.accent, clay.pillWarnText, clay.pillActiveText, clay.accentDark, clay.pillDangerText];

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Header */}
        <DmsPageHeader title="Reports & Compliance" subtitle="Live compliance summary and document analytics across your fleet.">
          <button
            onClick={() => {
              window.print();
            }}
            style={{ ...dmsBtn('primary'), padding: '0 18px', height: 36 }}
          >
            <Icon iconName="Print" aria-hidden="true" style={{ fontSize: 13 }} /> Print / Export Summary
          </button>
        </DmsPageHeader>

        {/* Date / Filter bar */}
        <div style={{ ...DMS_SECTION_CARD, padding: 14, flexDirection: 'row', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: 'var(--vdms-text-muted)' }}>
            <Icon iconName="Calendar" aria-hidden="true" style={{ fontSize: 13 }} />
            <span style={{ fontSize: 13, fontWeight: 600, color: 'var(--vdms-text)' }}>Live Fleet Data</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 'auto' }}>
            <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)', fontWeight: 600 }}>Filter by Vessel:</span>
            <select
              value={selectedVessel}
              onChange={e => host.setState({ reportsSelectedVessel: e.target.value })}
              style={{ border: '1px solid var(--vdms-line-strong)', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none', background: 'var(--vdms-surface)', color: 'var(--vdms-text)', fontWeight: 500 }}
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
          <div style={DMS_SECTION_CARD}>
            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Total Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: clay.accent, marginTop: 4 }}>{totalDocs}</div>
          </div>
          <div style={DMS_SECTION_CARD}>
            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Expired Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: expiredDocs > 0 ? dmsTone('danger').fg : 'var(--vdms-text)', marginTop: 4 }}>{expiredDocs}</div>
          </div>
          <div style={DMS_SECTION_CARD}>
            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Expiring in 30 days</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: expiringSoonDocs > 0 ? dmsTone('warning').fg : 'var(--vdms-text)', marginTop: 4 }}>{expiringSoonDocs}</div>
          </div>
          <div style={DMS_SECTION_CARD}>
            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Valid Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: dmsTone('success').fg, marginTop: 4 }}>{validDocs}</div>
          </div>
        </div>

        {/* Charts Grid */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          {/* Documents by Type */}
          <div style={DMS_SECTION_CARD}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: 'var(--vdms-text)' }}>Documents by Category / Type</h3>
            {typeList.length === 0 ? (
              <div style={{ padding: '32px 16px', textAlign: 'center', color: 'var(--vdms-text-faint)', fontSize: 13 }}>
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
                        <span style={{ color: 'var(--vdms-text-muted)', fontWeight: 600 }}>{tName}</span>
                        <span style={{ fontWeight: 700, color: 'var(--vdms-text)' }}>{tCount} ({pct}%)</span>
                      </div>
                      <div style={{ background: 'var(--vdms-surface-alt)', borderRadius: 4, height: 10, overflow: 'hidden' }}>
                        <div style={{ width: `${pct}%`, background: color, height: '100%', borderRadius: 4, transition: 'width 0.3s' }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Documents by Vessel (Bar Chart) */}
          <div style={DMS_SECTION_CARD}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: 'var(--vdms-text)' }}>Documents by Vessel</h3>
            {vesselList.length === 0 ? (
              <div style={{ padding: '32px 16px', textAlign: 'center', color: 'var(--vdms-text-faint)', fontSize: 13 }}>
                No vessel document data available yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {vesselList.map(([vName, vCount]) => (
                  <div key={vName} style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: 12 }}>
                    <span style={{ width: 110, color: 'var(--vdms-text-muted)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={vName}>
                      <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /> {vName}
                    </span>
                    <div style={{ flex: 1, background: 'var(--vdms-surface-alt)', borderRadius: 4, height: 16, overflow: 'hidden' }}>
                      <div style={{ width: `${(vCount / maxVesselDocs) * 100}%`, background: clay.accent, height: '100%', borderRadius: 4, transition: 'width 0.3s' }} />
                    </div>
                    <span style={{ width: 28, fontWeight: 700, textAlign: 'right', color: 'var(--vdms-text)' }}>{vCount}</span>
                  </div>
                ))}
                <div style={{ fontSize: 10, color: 'var(--vdms-text-faint)', textAlign: 'center', marginTop: 8 }}>No. of Documents</div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
}
