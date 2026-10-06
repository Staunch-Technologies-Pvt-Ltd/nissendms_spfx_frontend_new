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
import { Icon } from '@fluentui/react/lib/Icon';
import { clay } from '../clayTheme';
import {
  DmsPageHeader, DMS_TABLE_CARD, DMS_TABLE, DMS_TH, DMS_TR, DMS_TD,
  dmsBtn, dmsTone, dmsGrid,
} from '../dmsDesignSystem';

export function renderBentoEmailDashboardPage(host: VesselEmail): React.ReactElement {
    const {
      bentoLogs, panelLoading, bentoStatusFilter, bentoSearch,
      bentoComposeOpen, bentoComposeForm, bentoComposeBusy, bentoComposeMsg, bentoComposeErr,
      bentoDetailLog, vessels, bentoUploadFile, bentoUploadVessel, bentoUploadTag, bentoUploadBusy,
      bentoUploadMsg, bentoUploadErr, bentoClearAllBusy, bentoClearAllErr,
    } = host.state;

    const totalCount = bentoLogs.length;
    const completedCount = bentoLogs.filter(l => l.status === 'completed' || l.status === 'success' || l.status === 'sent').length;
    const pendingCount = bentoLogs.filter(l => l.status === 'pending').length;
    const failedCount = bentoLogs.filter(l => l.status === 'failed' || l.status === 'error').length;

    const q = bentoSearch.toLowerCase();
    const filtered = bentoLogs.filter(log => {
      const st = log.status === 'success' || log.status === 'sent' ? 'completed' : log.status;
      if (bentoStatusFilter !== 'all' && st !== bentoStatusFilter) return false;
      if (q) {
        return [
          log.vessel_name ?? '',
          log.subject ?? '',
          log.datasource_tag_used ?? '',
          log.recipient ?? '',
          ...(log.attachment_names || []),
        ].some(s => s.toLowerCase().indexOf(q) !== -1);
      }
      return true;
    });

    const completedTone = dmsTone('success');
    const pendingTone = dmsTone('warning');
    const failedTone = dmsTone('danger');
    const statCardStyle: React.CSSProperties = {
      background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-line)',
      padding: '14px 18px', boxShadow: clay.shadowRaised,
    };

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <DmsPageHeader
          title="AI Bento Email Dashboard"
          subtitle="Automated document tagging, status tracking, and Graph email dispatching."
        >
          <button
            style={dmsBtn('primary')}
            onClick={() => host.setState({
              bentoComposeOpen: true,
              bentoComposeMsg: null,
              bentoComposeErr: null,
              bentoComposeForm: {
                vessel_name: '',
                datasource_tag: 'mail',
                subject_text: host._buildAutoSubject('', 'mail', ''),
                body: '',
                file: null,
                existing_attachment: '',
                recipient: '',
              },
            })}
          >
            <Icon iconName="Mail" aria-hidden="true" style={{ fontSize: 13 }} /> Compose & Dispatch Email
          </button>
          <button
            disabled={bentoClearAllBusy || totalCount === 0}
            style={dmsBtn('danger', !(bentoClearAllBusy || totalCount === 0))}
            onClick={() => {
              if (!confirm(`Clear all ${totalCount} AI Bento Email log${totalCount === 1 ? '' : 's'}? This cannot be undone.`)) return;
              void host._clearAllBentoLogs();
            }}
          >
            {bentoClearAllBusy ? 'Clearing…' : <><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 13 }} /> Clear All</>}
          </button>
        </DmsPageHeader>

        {bentoClearAllErr && (
          <div style={{ background: dmsTone('danger').bg, border: '1px solid var(--vdms-line)', borderRadius: 8, padding: '8px 14px', fontSize: 12, color: dmsTone('danger').fg }}>
            {bentoClearAllErr}
          </div>
        )}

        <div style={dmsGrid(180)}>
          <div style={statCardStyle}>
            <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: 'var(--vdms-text-muted)' }}>Total Processed</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: 'var(--vdms-text)', marginTop: 4 }}>{totalCount}</div>
          </div>
          <div style={statCardStyle}>
            <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: completedTone.fg }}>Completed</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: completedTone.fg, marginTop: 4 }}>{completedCount}</div>
          </div>
          <div style={statCardStyle}>
            <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: pendingTone.fg }}>Pending</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: pendingTone.fg, marginTop: 4 }}>{pendingCount}</div>
          </div>
          <div style={statCardStyle}>
            <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: failedTone.fg }}>Failed</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: failedTone.fg, marginTop: 4 }}>{failedCount}</div>
          </div>
        </div>

        {/* Logs Table */}
        <div style={{ ...DMS_TABLE_CARD, overflowX: 'auto' }}>
          <table style={{ ...DMS_TABLE, minWidth: 900 }}>
            <thead>
              <tr>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>#</th>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>Vessel</th>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>Tag</th>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>Subject</th>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>Recipient</th>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>Status</th>
                <th style={{ ...DMS_TH, whiteSpace: 'nowrap' }}>Attached File</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((log) => (
                <tr key={log.id} style={DMS_TR}>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>#{log.id}</td>
                  <td style={{ ...DMS_TD, fontWeight: 600, color: 'var(--vdms-text)', whiteSpace: 'nowrap' }}>{log.vessel_name || '—'}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>{badge('blue', log.datasource_tag_used)}</td>
                  <td style={{ ...DMS_TD, color: 'var(--vdms-text)', minWidth: 200 }}>{log.subject}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>{log.recipient}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>
                    {badge(log.status === 'completed' || log.status === 'success' ? 'green' : log.status === 'pending' ? 'orange' : 'red', log.status)}
                  </td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>
                    {log.attachment_names && log.attachment_names.length > 0
                      ? log.attachment_names.map((name, i) => (
                          <span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: dmsTone('accent').bg, border: '1px solid var(--vdms-line)', borderRadius: 6, padding: '2px 8px', fontSize: 11, color: dmsTone('accent').fg, fontWeight: 600, marginRight: 4 }}>
                            <Icon iconName="Attach" aria-hidden="true" style={{ fontSize: 11 }} /> {name}
                          </span>
                        ))
                      : <span style={{ color: 'var(--vdms-text-faint)', fontSize: 11 }}>—</span>
                    }
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
}
