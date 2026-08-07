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

export function renderReportsPage(host: VesselEmail): React.ReactElement {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Reports</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Generate and view analytics reports.</p>
          </div>

          <button
            onClick={() => alert('Downloading report PDF...')}
            style={{
              background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
              padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              display: 'inline-flex', alignItems: 'center', gap: 6,
            }}
          >
            ⬇ Download Report
          </button>
        </div>

        {/* Date / Filter bar */}
        <div style={{ background: '#fff', borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', display: 'flex', gap: 12, alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: '#334155' }}>
            <span>📅</span>
            <input type="text" defaultValue="May 1, 2024 - May 31, 2024" style={{ border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none' }} />
          </div>
          <select style={{ border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none', background: '#fff' }}>
            <option>All Vessels</option>
            <option>Ocean Star</option>
            <option>Sea Breeze</option>
          </select>
        </div>

        {/* 4 Stat Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16 }}>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Total Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#0078d4', marginTop: 4 }}>156</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Expired Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#ef4444', marginTop: 4 }}>8</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Expiring in 30 days</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#f59e0b', marginTop: 4 }}>23</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Valid Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#10b981', marginTop: 4 }}>125</div>
          </div>
        </div>

        {/* Charts Grid */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          {/* Documents by Type (Donut Chart) */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Documents by Type</h3>
            <div style={{ display: 'flex', alignItems: 'center', gap: 24 }}>
              <div style={{ position: 'relative', width: 140, height: 140 }}>
                <svg width="140" height="140" viewBox="0 0 42 42">
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#e2e8f0" strokeWidth="5" />
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#2563eb" strokeWidth="5" strokeDasharray="29 71" strokeDashoffset="25" />
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f59e0b" strokeWidth="5" strokeDasharray="26 74" strokeDashoffset="96" />
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#10b981" strokeWidth="5" strokeDasharray="22 78" strokeDashoffset="70" />
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#8b5cf6" strokeWidth="5" strokeDasharray="13 87" strokeDashoffset="48" />
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#06b6d4" strokeWidth="5" strokeDasharray="10 90" strokeDashoffset="35" />
                </svg>
              </div>

              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 6, fontSize: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Insurance</span>
                  <span style={{ fontWeight: 600 }}>45 (29%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Crew</span>
                  <span style={{ fontWeight: 600 }}>40 (26%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Maintenance</span>
                  <span style={{ fontWeight: 600 }}>35 (22%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Certificate</span>
                  <span style={{ fontWeight: 600 }}>20 (13%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Survey</span>
                  <span style={{ fontWeight: 600 }}>16 (10%)</span>
                </div>
                <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 6, marginTop: 4, display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}>
                  <span>Total</span>
                  <span>156</span>
                </div>
              </div>
            </div>
          </div>

          {/* Documents by Vessel (Bar Chart) */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Documents by Vessel</h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              {[
                { name: 'Ocean Star', count: 45, max: 50 },
                { name: 'Sea Breeze', count: 32, max: 50 },
                { name: 'Blue Horizon', count: 28, max: 50 },
                { name: 'Pacific Dawn', count: 26, max: 50 },
                { name: 'Atlantic Wave', count: 25, max: 50 },
              ].map(v => (
                <div key={v.name} style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: 12 }}>
                  <span style={{ width: 90, color: '#475569', fontWeight: 500 }}>{v.name}</span>
                  <div style={{ flex: 1, background: '#f1f5f9', borderRadius: 4, height: 16, overflow: 'hidden' }}>
                    <div style={{ width: `${(v.count / v.max) * 100}%`, background: '#2563eb', height: '100%', borderRadius: 4 }} />
                  </div>
                  <span style={{ width: 24, fontWeight: 600, textAlign: 'right' }}>{v.count}</span>
                </div>
              ))}
              <div style={{ fontSize: 10, color: '#94a3b8', textAlign: 'center', marginTop: 8 }}>No. of Documents</div>
            </div>
          </div>
        </div>
      </div>
    );
}
