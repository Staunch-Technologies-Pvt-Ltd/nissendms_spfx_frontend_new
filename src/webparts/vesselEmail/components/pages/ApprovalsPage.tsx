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

export function renderApprovalsPage(host: VesselEmail): React.ReactElement {
    const { approvalsList, approvalTab, panelLoading } = host.state;
    const filtered = approvalsList.filter(a => a.status === approvalTab);

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Approvals</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Review and take action on pending approvals.</p>
        </div>

        {/* Tabs */}
        <div style={{ display: 'flex', gap: 8, borderBottom: '1px solid #e2e8f0', paddingBottom: 8 }}>
          {(['Pending', 'Approved', 'Rejected'] as const).map(tab => (
            <button
              key={tab}
              onClick={() => host.setState({ approvalTab: tab })}
              style={{
                border: 'none', background: approvalTab === tab ? '#eff6ff' : 'transparent',
                color: approvalTab === tab ? '#0078d4' : '#64748b', fontWeight: approvalTab === tab ? 700 : 500,
                fontSize: 13, padding: '6px 14px', borderRadius: 6, cursor: 'pointer',
              }}
            >
              {tab} ({approvalsList.filter(a => a.status === tab).length})
            </button>
          ))}
        </div>

        {/* Table */}
        {panelLoading ? (
          <div style={{ padding: 32, textAlign: 'center', color: '#64748b', fontSize: 13 }}>⏳ Loading approvals...</div>
        ) : (
        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflowX: 'auto' }}>
          <table style={{ width: '100%', minWidth: 700, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Document Name</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Vessel</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Requested By</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Requested On</th>
                <th style={{ padding: '10px 16px', textAlign: 'right', whiteSpace: 'nowrap' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={5} style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No {approvalTab.toLowerCase()} approvals.</td></tr>
              ) : filtered.map(app => (
                <tr key={app.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={app.documentName}>{app.documentName}</td>
                  <td style={{ padding: '12px 16px', color: '#475569', whiteSpace: 'nowrap' }}>{app.vessel}</td>
                  <td style={{ padding: '12px 16px', color: '#475569', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={app.requestedBy}>{app.requestedBy}</td>
                  <td style={{ padding: '12px 16px', color: '#64748b', whiteSpace: 'nowrap' }}>{app.requestedOn}</td>
                  <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                    {app.status === 'Pending' ? (
                      <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                        <button
                          onClick={async () => {
                            const userEmail = host.props.userEmail || '';
                            const adminParam = userEmail ? `?admin=${encodeURIComponent(userEmail)}` : '';
                            try {
                              await fetch(`${host._base()}/api/approvals/${app.id}/approve${adminParam}`, { method: 'POST', headers: host._headers() });
                            } catch { /* fallback to local update */ }
                            const updated = approvalsList.map(item => item.id === app.id ? { ...item, status: 'Approved' as const } : item);
                            host.setState({ approvalsList: updated });
                            host._filesLoadedForFolders.clear();
                            host._filesLoadedForVessels.clear();
                            host._loadData();
                            // Re-fetch files for the approved item's vessel so list view updates
                            if (app.vessel && app.vessel !== '—') {
                              const vesselName = app.vessel;
                              setTimeout(() => {
                                host._filesLoadedForVessels.add(vesselName);
                                host._loadFilesForVessel(vesselName).catch(() => undefined);
                              }, 2000);
                            }
                          }}
                          style={{ border: 'none', background: '#dff6dd', color: '#107c10', width: 28, height: 28, borderRadius: 4, cursor: 'pointer', fontWeight: 700 }}
                          title="Approve"
                        >
                          ✓
                        </button>
                        <button
                          onClick={async () => {
                            const userEmail = host.props.userEmail || '';
                            const adminParam = userEmail ? `?admin=${encodeURIComponent(userEmail)}` : '';
                            try {
                              await fetch(`${host._base()}/api/approvals/${app.id}/reject${adminParam}`, {
                                method: 'POST',
                                headers: { ...host._headers(), 'Content-Type': 'application/json' },
                                body: JSON.stringify({ reason: null }),
                              });
                            } catch { /* fallback to local update */ }
                            const updated = approvalsList.map(item => item.id === app.id ? { ...item, status: 'Rejected' as const } : item);
                            host.setState({ approvalsList: updated });
                          }}
                          style={{ border: 'none', background: '#fde7e9', color: '#a4262c', width: 28, height: 28, borderRadius: 4, cursor: 'pointer', fontWeight: 700 }}
                          title="Reject"
                        >
                          ✕
                        </button>
                      </div>
                    ) : (
                      badge(app.status === 'Approved' ? 'green' : 'red', app.status)
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        )}
      </div>
    );
}
