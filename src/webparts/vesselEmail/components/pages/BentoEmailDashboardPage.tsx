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

export function renderBentoEmailDashboardPage(host: VesselEmail): React.ReactElement {
    const {
      bentoLogs, panelLoading, bentoStatusFilter, bentoSearch,
      bentoComposeOpen, bentoComposeForm, bentoComposeBusy, bentoComposeMsg, bentoComposeErr,
      bentoDetailLog, vessels, bentoUploadFile, bentoUploadVessel, bentoUploadTag, bentoUploadBusy,
      bentoUploadMsg, bentoUploadErr,
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

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)', borderRadius: 12, padding: '24px 28px', color: '#fff', boxShadow: '0 4px 16px rgba(0,0,0,0.15)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 }}>
            <div>
              <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: 'rgba(255,255,255,0.1)', padding: '4px 12px', borderRadius: 20, fontSize: 12, fontWeight: 600, color: '#38bdf8', marginBottom: 8 }}>
                🤖 AI BENTO AUTOMATION ENGINE
              </div>
              <h1 style={{ margin: 0, fontSize: 22, fontWeight: 800 }}>AI Bento Email Dashboard</h1>
              <p style={{ margin: '6px 0 0', fontSize: 13, color: '#94a3b8', maxWidth: 650 }}>
                Automated document tagging, status tracking, and Graph email dispatching.
              </p>
            </div>
            <button
              style={{ background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 22px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}
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
              ✉ Compose & Dispatch Email
            </button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 14, marginTop: 22 }}>
            <div style={{ background: 'rgba(255,255,255,0.06)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(255,255,255,0.1)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#94a3b8' }}>Total Processed</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#fff', marginTop: 4 }}>{totalCount}</div>
            </div>
            <div style={{ background: 'rgba(16, 185, 129, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(16, 185, 129, 0.25)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#6ee7b7' }}>Completed</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#34d399', marginTop: 4 }}>{completedCount}</div>
            </div>
            <div style={{ background: 'rgba(245, 158, 11, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(245, 158, 11, 0.25)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#fcd34d' }}>Pending</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#fbbf24', marginTop: 4 }}>{pendingCount}</div>
            </div>
            <div style={{ background: 'rgba(239, 68, 68, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(239, 68, 68, 0.25)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#fca5a5' }}>Failed</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#f87171', marginTop: 4 }}>{failedCount}</div>
            </div>
          </div>
        </div>

        {/* Logs Table */}
        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflowX: 'auto' }}>
          <table style={{ width: '100%', minWidth: 900, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left', whiteSpace: 'nowrap' }}>
                <th style={{ padding: '10px 16px' }}>#</th>
                <th style={{ padding: '10px 16px' }}>Vessel</th>
                <th style={{ padding: '10px 16px' }}>Tag</th>
                <th style={{ padding: '10px 16px' }}>Subject</th>
                <th style={{ padding: '10px 16px' }}>Recipient</th>
                <th style={{ padding: '10px 16px' }}>Status</th>
                <th style={{ padding: '10px 16px' }}>Attached File</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((log) => (
                <tr key={log.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '10px 16px', color: '#64748b', whiteSpace: 'nowrap' }}>#{log.id}</td>
                  <td style={{ padding: '10px 16px', fontWeight: 600, whiteSpace: 'nowrap' }}>{log.vessel_name || '—'}</td>
                  <td style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>{badge('blue', log.datasource_tag_used)}</td>
                  <td style={{ padding: '10px 16px', color: '#1e293b', minWidth: 200 }}>{log.subject}</td>
                  <td style={{ padding: '10px 16px', color: '#475569', whiteSpace: 'nowrap' }}>{log.recipient}</td>
                  <td style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>
                    {badge(log.status === 'completed' || log.status === 'success' ? 'green' : log.status === 'pending' ? 'orange' : 'red', log.status)}
                  </td>
                  <td style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>
                    {log.attachment_names && log.attachment_names.length > 0
                      ? log.attachment_names.map((name, i) => (
                          <span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 6, padding: '2px 8px', fontSize: 11, color: '#0284c7', fontWeight: 600, marginRight: 4 }}>
                            📎 {name}
                          </span>
                        ))
                      : <span style={{ color: '#94a3b8', fontSize: 11 }}>—</span>
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
