import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { Icon } from '@fluentui/react/lib/Icon';
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
import { clay } from '../clayTheme';
import { DmsPageHeader, DMS_TABLE_CARD, dmsBtn, dmsTone } from '../dmsDesignSystem';

// Map action_type values to human-readable labels and badge colours — tone
// keys resolve through the shared `dmsTone()` palette so these stay in sync
// with the rest of the app's semantic colors.
const ACTION_LABELS: Record<string, { label: string; tone: 'accent' | 'warning' | 'danger' | 'success' | 'neutral' }> = {
  upload:          { label: 'Upload',         tone: 'accent' },
  delete_file:     { label: 'Delete File',    tone: 'warning' },
  delete_folder:   { label: 'Delete Folder',  tone: 'warning' },
  delete_vessel:   { label: 'Delete Vessel',  tone: 'danger' },
  create_vessel:   { label: 'New Vessel',     tone: 'success' },
  create_folder:   { label: 'New Folder',     tone: 'accent' },
  update_vessel:   { label: 'Edit Vessel',    tone: 'accent' },
  rename_folder:   { label: 'Rename Folder',  tone: 'accent' },
  move_file:       { label: 'Move File',      tone: 'success' },
};

function actionBadge(actionType?: string): React.ReactElement | null {
  if (!actionType) return null;
  const meta = ACTION_LABELS[actionType] || { label: actionType.replace(/_/g, ' '), tone: 'neutral' as const };
  const { bg, fg } = dmsTone(meta.tone);
  const cfg = { label: meta.label, color: fg, bg };
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
      { key: 'Pending',  count: pendingCount,  color: dmsTone('warning').fg, bg: dmsTone('warning').bg },
      { key: 'Approved', count: approvedCount, color: dmsTone('success').fg, bg: dmsTone('success').bg },
      { key: 'Rejected', count: rejectedCount, color: dmsTone('danger').fg, bg: dmsTone('danger').bg },
    ];

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <DmsPageHeader title="Approvals" subtitle="Review and take action on pending approvals." />

        {/* Tabs + Clear All button */}
        <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', borderBottom: '2px solid var(--vdms-line)', paddingBottom: 0 }}>
          <div style={{ display: 'flex', gap: 8 }}>
          {tabMeta.map(({ key, count, color, bg }) => (
            <button
              key={key}
              onClick={() => host.setState({ approvalTab: key })}
              style={{
                border: 'none',
                borderBottom: approvalTab === key ? `2px solid ${color}` : '2px solid transparent',
                background: 'transparent',
                color: approvalTab === key ? color : 'var(--vdms-text-muted)',
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
                background: approvalTab === key ? bg : 'var(--vdms-surface-alt)',
                color: approvalTab === key ? color : 'var(--vdms-text-faint)',
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
              style={{ ...dmsBtn('danger'), padding: '5px 12px', marginBottom: 4 }}
              title={`Remove all ${approvalTab.toLowerCase()} entries`}
            >
              <Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 12 }} /> Clear All {approvalTab}
            </button>
          )}
        </div>

        {/* Table */}
        {panelLoading ? (
          <div style={{ padding: 48, textAlign: 'center', color: 'var(--vdms-text-muted)', fontSize: 13 }}>
            <div style={{ fontSize: 28, marginBottom: 8 }}><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 28 }} /></div>
            Loading approvals...
          </div>
        ) : filtered.length === 0 ? (
          <div style={{
            padding: 48, textAlign: 'center', color: 'var(--vdms-text-faint)', fontSize: 14,
            background: 'var(--vdms-surface-alt)', borderRadius: 12, border: '1px dashed var(--vdms-line)',
          }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>
              {approvalTab === 'Pending' ? <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 32 }} /> : approvalTab === 'Approved' ? <Icon iconName="Completed" aria-hidden="true" style={{ fontSize: 32 }} /> : <Icon iconName="Blocked2" aria-hidden="true" style={{ fontSize: 32 }} />}
            </div>
            <div style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>No {approvalTab.toLowerCase()} approvals</div>
            {approvalTab === 'Pending' && (
              <div style={{ fontSize: 12, marginTop: 4, color: 'var(--vdms-text-faint)' }}>All caught up! No pending actions waiting for review.</div>
            )}
          </div>
        ) : (
        <div style={DMS_TABLE_CARD}>
          <table style={{ width: '100%', minWidth: 700, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: 'var(--vdms-surface-alt)', borderBottom: '2px solid var(--vdms-line)', color: 'var(--vdms-text-muted)', fontSize: 11, textTransform: 'uppercase', textAlign: 'left', letterSpacing: '0.04em' }}>
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
                    borderBottom: '1px solid var(--vdms-border-soft)',
                    background: idx % 2 === 0 ? 'var(--vdms-surface)' : 'var(--vdms-surface-alt)',
                    transition: 'background 0.12s',
                  }}
                  onMouseEnter={e => (e.currentTarget.style.background = clay.accentSoft)}
                  onMouseLeave={e => (e.currentTarget.style.background = idx % 2 === 0 ? 'var(--vdms-surface)' : 'var(--vdms-surface-alt)')}
                >
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--vdms-text)', maxWidth: 260 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 0, flexWrap: 'wrap' }}>
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 200 }} title={app.documentName}>
                        {app.documentName}
                      </span>
                      {actionBadge(app.actionType)}
                    </div>
                  </td>
                  <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)', whiteSpace: 'nowrap' }}>
                    {app.vessel && app.vessel !== '—' ? (
                      <span style={{
                        display: 'inline-flex', alignItems: 'center', gap: 4,
                        background: clay.accentSoft, color: clay.accent, borderRadius: 6,
                        padding: '2px 8px', fontSize: 12, fontWeight: 600,
                      }}>
                        <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /> {app.vessel}
                      </span>
                    ) : (
                      <span style={{ color: 'var(--vdms-text-faint)' }}>—</span>
                    )}
                  </td>
                  <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={app.requestedBy}>
                    {app.requestedBy}
                  </td>
                  <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)', whiteSpace: 'nowrap', fontSize: 12 }}>
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
                            border: 'none', background: dmsTone('success').bg, color: dmsTone('success').fg,
                            width: 32, height: 32, borderRadius: 6, cursor: 'pointer', fontWeight: 700,
                            fontSize: 15, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            transition: 'all 0.15s',
                          }}
                          title="Approve"
                        >
                          <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 15 }} />
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
                            border: 'none', background: dmsTone('danger').bg, color: dmsTone('danger').fg,
                            width: 32, height: 32, borderRadius: 6, cursor: 'pointer', fontWeight: 700,
                            fontSize: 15, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            transition: 'all 0.15s',
                          }}
                          title="Reject"
                        >
                          <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 15 }} />
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
                            border: '1px solid var(--vdms-line)', background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text-faint)',
                            width: 28, height: 28, borderRadius: 6, cursor: 'pointer', fontSize: 14,
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            transition: 'all 0.15s',
                          }}
                          title="Remove this entry"
                          onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = dmsTone('danger').bg; (e.currentTarget as HTMLButtonElement).style.color = dmsTone('danger').fg; }}
                          onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--vdms-surface-alt)'; (e.currentTarget as HTMLButtonElement).style.color = 'var(--vdms-text-faint)'; }}
                        >
                          <Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 14 }} />
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
