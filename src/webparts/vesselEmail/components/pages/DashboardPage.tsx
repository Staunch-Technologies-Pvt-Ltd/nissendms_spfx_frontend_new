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

export function renderDashboard(host: VesselEmail): React.ReactElement {
    const { rows, vessels, documentsList } = host.state;
    const userDisplayName = host.props.userDisplayName || 'Priya';

    const totalDocs = rows.length > 0 ? rows.length : 156;
    const expiringSoon = 23;
    const pendingApprovals = 7;
    const totalVessels = vessels.length;

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Welcome Banner */}
        <div style={{
          background: 'linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)',
          borderRadius: 12, padding: '24px 32px', display: 'flex', justifyContent: 'space-between',
          alignItems: 'center', border: '1px solid #bfdbfe', position: 'relative', overflow: 'hidden',
        }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#1e3a8a' }}>
              Welcome back, {userDisplayName}! 👋
            </h1>
            <p style={{ margin: '6px 0 0', fontSize: 13, color: '#3b82f6' }}>
              Here's what's happening with your vessels and documents.
            </p>
          </div>
          {/* Real Vessel Graphic Banner */}
          <div style={{ height: 85, width: 200, borderRadius: 8, overflow: 'hidden', flexShrink: 0, boxShadow: '0 4px 12px rgba(37,99,235,0.15)' }}>
            <img
              src={resolveImgUrl(null)}
              alt="Vessel Illustration"
              style={{ width: '100%', height: '100%', objectFit: 'cover' }}
            />
          </div>
        </div>

        {/* 4 Stat Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16 }}>
          {/* Stat 1 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>{totalDocs}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Total Documents</div>
              <button onClick={() => host._goToView('list')} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#f0f9ff', color: '#0284c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>📄</div>
          </div>

          {/* Stat 2 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>{expiringSoon}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Expiring Soon</div>
              <button onClick={() => host._goToView('list')} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#ecfdf5', color: '#10b981', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>✅</div>
          </div>

          {/* Stat 3 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>0{pendingApprovals}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Pending Approvals</div>
              <button onClick={() => host._goToView('approvals')} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#fffbeb', color: '#f59e0b', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>⏰</div>
          </div>

          {/* Stat 4 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>{totalVessels}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Total Vessels</div>
              <button onClick={() => host._goToView('vessels')} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#eff6ff', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>🚢</div>
          </div>
        </div>

        {/* Main Dashboard Grid: Recent Docs + Expiry Overview */}
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 20 }}>
          {/* Recent Documents Table */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Recent Documents</h3>
              <button onClick={() => host._goToView('list')} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>View all</button>
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: '1px solid #e2e8f0', textTransform: 'uppercase', fontSize: 11, color: '#64748b', textAlign: 'left' }}>
                  <th style={{ padding: '8px 12px' }}>Name</th>
                  <th style={{ padding: '8px 12px' }}>Vessel</th>
                  <th style={{ padding: '8px 12px' }}>Type</th>
                  <th style={{ padding: '8px 12px' }}>Modified</th>
                </tr>
              </thead>
              <tbody>
                {documentsList.slice(3, 7).map((doc, idx) => (
                  <tr key={doc.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ color: doc.name.endsWith('.pdf') ? '#ef4444' : doc.name.endsWith('.docx') ? '#2563eb' : '#10b981' }}>📄</span>
                      {doc.name}
                    </td>
                    <td style={{ padding: '10px 12px', color: '#475569' }}>{doc.vessel}</td>
                    <td style={{ padding: '10px 12px' }}>
                      {badge(doc.type === 'Insurance' ? 'purple' : doc.type === 'Crew' ? 'blue' : doc.type === 'Maintenance' ? 'green' : 'orange', doc.type)}
                    </td>
                    <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 12 }}>{doc.modified}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Document Expiry Overview Chart */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Document Expiry Overview</h3>
              <button onClick={() => host._goToView('reports')} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>View all</button>
            </div>

            {/* SVG Donut Chart */}
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16 }}>
              <div style={{ position: 'relative', width: 140, height: 140 }}>
                <svg width="140" height="140" viewBox="0 0 42 42">
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#e2e8f0" strokeWidth="5" />
                  {/* Valid 67% (Green) */}
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#10b981" strokeWidth="5" strokeDasharray="67 33" strokeDashoffset="25" />
                  {/* Expiring in 30 days 24% (Orange) */}
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f59e0b" strokeWidth="5" strokeDasharray="24 76" strokeDashoffset="58" />
                  {/* Expired 9% (Red) */}
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#ef4444" strokeWidth="5" strokeDasharray="9 91" strokeDashoffset="34" />
                </svg>
                <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
                  <span style={{ fontSize: 20, fontWeight: 800, color: '#0f172a' }}>96</span>
                  <span style={{ fontSize: 10, color: '#64748b', textTransform: 'uppercase' }}>Total</span>
                </div>
              </div>

              {/* Legend */}
              <div style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: 8, fontSize: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#ef4444' }} />
                    <span style={{ color: '#475569' }}>Expired</span>
                  </div>
                  <span style={{ fontWeight: 600, color: '#0f172a' }}>8 (9%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#f59e0b' }} />
                    <span style={{ color: '#475569' }}>Expiring in 30 days</span>
                  </div>
                  <span style={{ fontWeight: 600, color: '#0f172a' }}>23 (24%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#10b981' }} />
                    <span style={{ color: '#475569' }}>Valid</span>
                  </div>
                  <span style={{ fontWeight: 600, color: '#0f172a' }}>65 (67%)</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
}
