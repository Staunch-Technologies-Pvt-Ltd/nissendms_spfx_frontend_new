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

// Map action_type values to human-readable labels and badge colours
const ACTION_LABELS: Record<string, { label: string; color: string; bg: string }> = {
  upload:          { label: 'Upload',         color: '#0369a1', bg: '#e0f2fe' },
  delete_file:     { label: 'Delete File',    color: '#9a3412', bg: '#ffedd5' },
  delete_folder:   { label: 'Delete Folder',  color: '#7c2d12', bg: '#fef3c7' },
  delete_vessel:   { label: 'Delete Vessel',  color: '#991b1b', bg: '#fee2e2' },
  create_vessel:   { label: 'New Vessel',     color: '#065f46', bg: '#d1fae5' },
  create_folder:   { label: 'New Folder',     color: '#1e40af', bg: '#dbeafe' },
  update_vessel:   { label: 'Edit Vessel',    color: '#4338ca', bg: '#ede9fe' },
  rename_folder:   { label: 'Rename Folder',  color: '#5b21b6', bg: '#ede9fe' },
  move_file:       { label: 'Move File',      color: '#0f766e', bg: '#ccfbf1' },
};

function actionBadge(actionType?: string): React.ReactElement | null {
  if (!actionType) return null;
  const cfg = ACTION_LABELS[actionType] || { label: actionType.replace(/_/g, ' '), color: '#475569', bg: '#f1f5f9' };
  return (
    <span style={{
      display: 'inline-block', padding: '2px 8px', borderRadius: 20, fontSize: 10,
      fontWeight: 700, background: cfg.bg, color: cfg.color,
      whiteSpace: 'nowrap', marginLeft: 6, verticalAlign: 'middle',
      textTransform: 'uppercase', letterSpacing: '0.03em',
    }}>
      {cfg.label}
    </span>
  );
}

export function renderApprovalsPage(host: VesselEmail): React.ReactElement {
    const { approvalsList, approvalTab, panelLoading } = host.state;
    const filtered = approvalsList.filter(a => a.status === approvalTab);

    const pendingCount   = approvalsList.filter(a => a.status === 'Pending').length;
    const approvedCount  = approvalsList.filter(a => a.status === 'Approved').length;
    const rejectedCount  = approvalsList.filter(a => a.status === 'Rejected').length;

    const tabMeta: { key: 'Pending' | 'Approved' | 'Rejected'; count: number; color: string; bg: string }[] = [
      { key: 'Pending',  count: pendingCount,  color: '#b45309', bg: '#fef3c7' },
      { key: 'Approved', count: approvedCount, color: '#065f46', bg: '#d1fae5' },
      { key: 'Rejected', count: rejectedCount, color: '#991b1b', bg: '#fee2e2' },
    ];

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Approvals</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Review and take action on pending approvals.</p>
        </div>

        {/* Tabs + Clear All button */}
        <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', borderBottom: '2px solid #e2e8f0', paddingBottom: 0 }}>
          <div style={{ display: 'flex', gap: 8 }}>
          {tabMeta.map(({ key, count, color, bg }) => (
            <button
              key={key}
              onClick={() => host.setState({ approvalTab: key })}
              style={{
                border: 'none',
                borderBottom: approvalTab === key ? `2px solid ${color}` : '2px solid transparent',
                background: 'transparent',
                color: approvalTab === key ? color : '#64748b',
                fontWeight: approvalTab === key ? 700 : 500,
                fontSize: 13,
                padding: '8px 16px',
                borderRadius: 0,
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                marginBottom: -2,
                transition: 'all 0.15s',
              }}
            >
              {key}
              <span style={{
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                minWidth: 20, height: 20, padding: '0 6px',
                borderRadius: 20, fontSize: 10, fontWeight: 700,
                background: approvalTab === key ? bg : '#f1f5f9',
                color: approvalTab === key ? color : '#94a3b8',
              }}>
                {count}
              </span>
            </button>
          ))}
          </div>
          {/* Clear All button — only for Approved / Rejected tabs with items */}
          {approvalTab !== 'Pending' && filtered.length > 0 && (
            <button
              onClick={async () => {
                if (!window.confirm(`Remove all ${filtered.length} ${approvalTab.toLowerCase()} entries? This cannot be undone.`)) return;
                const userEmail = host.props.userEmail || '';
                const adminParam = userEmail ? `?admin=${encodeURIComponent(userEmail)}` : '';
                try {
                  await fetch(`${host._base()}/api/approvals?status=${approvalTab.toLowerCase()}${adminParam.replace('?', '&')}`, {
                    method: 'DELETE', headers: host._headers(),
                  });
                } catch { /* offline fallback */ }
                host.setState({ approvalsList: host.state.approvalsList.filter(a => a.status !== approvalTab) });
              }}
              style={{
                border: '1px solid #fca5a5', background: '#fef2f2', color: '#dc2626',
                borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600,
                cursor: 'pointer', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 5,
              }}
              title={`Remove all ${approvalTab.toLowerCase()} entries`}
            >
              🗑 Clear All {approvalTab}
            </button>
          )}
        </div>

        {/* Table */}
        {panelLoading ? (
          <div style={{ padding: 48, textAlign: 'center', color: '#64748b', fontSize: 13 }}>
            <div style={{ fontSize: 28, marginBottom: 8 }}>⏳</div>
            Loading approvals...
          </div>
        ) : filtered.length === 0 ? (
          <div style={{
            padding: 48, textAlign: 'center', color: '#94a3b8', fontSize: 14,
            background: '#f8fafc', borderRadius: 12, border: '1px dashed #e2e8f0',
          }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>
              {approvalTab === 'Pending' ? '✅' : approvalTab === 'Approved' ? '🎉' : '🚫'}
            </div>
            <div style={{ fontWeight: 600, color: '#64748b' }}>No {approvalTab.toLowerCase()} approvals</div>
            {approvalTab === 'Pending' && (
              <div style={{ fontSize: 12, marginTop: 4, color: '#94a3b8' }}>All caught up! No pending actions waiting for review.</div>
            )}
          </div>
        ) : (
        <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', overflow: 'hidden', boxShadow: '0 1px 4px rgba(0,0,0,0.04)' }}>
          <table style={{ width: '100%', minWidth: 700, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left', letterSpacing: '0.04em' }}>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Document / Action</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Vessel</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Requested By</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Requested On</th>
                <th style={{ padding: '10px 16px', textAlign: 'right', whiteSpace: 'nowrap' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((app, idx) => (
                <tr
                  key={app.id}
                  style={{
                    borderBottom: '1px solid #f1f5f9',
                    background: idx % 2 === 0 ? '#fff' : '#fafafa',
                    transition: 'background 0.12s',
                  }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f0f9ff')}
                  onMouseLeave={e => (e.currentTarget.style.background = idx % 2 === 0 ? '#fff' : '#fafafa')}
                >
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', maxWidth: 260 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 0, flexWrap: 'wrap' }}>
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 200 }} title={app.documentName}>
                        {app.documentName}
                      </span>
                      {actionBadge(app.actionType)}
                    </div>
                  </td>
                  <td style={{ padding: '12px 16px', color: '#475569', whiteSpace: 'nowrap' }}>
                    {app.vessel && app.vessel !== '—' ? (
                      <span style={{
                        display: 'inline-flex', alignItems: 'center', gap: 4,
                        background: '#eff6ff', color: '#1d4ed8', borderRadius: 6,
                        padding: '2px 8px', fontSize: 12, fontWeight: 600,
                      }}>
                        🚢 {app.vessel}
                      </span>
                    ) : (
                      <span style={{ color: '#cbd5e1' }}>—</span>
                    )}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#475569', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={app.requestedBy}>
                    {app.requestedBy}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#64748b', whiteSpace: 'nowrap', fontSize: 12 }}>
                    {app.requestedOn}
                  </td>
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
                            if (app.vessel && app.vessel !== '—') {
                              const vesselName = app.vessel;
                              setTimeout(() => {
                                host._filesLoadedForVessels.add(vesselName);
                                host._loadFilesForVessel(vesselName).catch(() => undefined);
                              }, 2000);
                            }
                          }}
                          style={{
                            border: 'none', background: '#dff6dd', color: '#107c10',
                            width: 32, height: 32, borderRadius: 6, cursor: 'pointer', fontWeight: 700,
                            fontSize: 15, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            transition: 'all 0.15s',
                          }}
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
                                body: JSON.stringify({ reason: 'Rejected by admin' }),
                              });
                            } catch { /* fallback to local update */ }
                            const updated = approvalsList.map(item => item.id === app.id ? { ...item, status: 'Rejected' as const } : item);
                            host.setState({ approvalsList: updated });
                          }}
                          style={{
                            border: 'none', background: '#fde7e9', color: '#a4262c',
                            width: 32, height: 32, borderRadius: 6, cursor: 'pointer', fontWeight: 700,
                            fontSize: 15, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            transition: 'all 0.15s',
                          }}
                          title="Reject"
                        >
                          ✕
                        </button>
                      </div>
                    ) : (
                      <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end', alignItems: 'center' }}>
                        {badge(app.status === 'Approved' ? 'green' : 'red', app.status)}
                        <button
                          onClick={async () => {
                            const userEmail = host.props.userEmail || '';
                            const adminParam = userEmail ? `?admin=${encodeURIComponent(userEmail)}` : '';
                            try {
                              await fetch(`${host._base()}/api/approvals/${app.id}${adminParam}`, {
                                method: 'DELETE', headers: host._headers(),
                              });
                            } catch { /* offline fallback */ }
                            host.setState({ approvalsList: host.state.approvalsList.filter(a => a.id !== app.id) });
                          }}
                          style={{
                            border: '1px solid #e2e8f0', background: '#f8fafc', color: '#94a3b8',
                            width: 28, height: 28, borderRadius: 6, cursor: 'pointer', fontSize: 14,
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            transition: 'all 0.15s',
                          }}
                          title="Remove this entry"
                          onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = '#fee2e2'; (e.currentTarget as HTMLButtonElement).style.color = '#dc2626'; }}
                          onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = '#f8fafc'; (e.currentTarget as HTMLButtonElement).style.color = '#94a3b8'; }}
                        >
                          🗑
                        </button>
                      </div>
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
