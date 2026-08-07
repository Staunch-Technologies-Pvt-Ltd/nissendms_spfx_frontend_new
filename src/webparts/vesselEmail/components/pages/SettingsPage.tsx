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

export function renderSettingsPage(host: VesselEmail): React.ReactElement {
    const { settingsTab, settingsForm, settingsSavedMsg } = host.state;

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Settings</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Configure application settings and preferences.</p>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '200px 1fr', gap: 24, background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 24 }}>
          {/* Settings Left Nav */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4, borderRight: '1px solid #f1f5f9', paddingRight: 16 }}>
            {[
              'General', 'Document Settings', 'Notification Settings',
              'Permission Settings', 'Integration', 'Audit Logs'
            ].map(tab => (
              <button
                key={tab}
                onClick={() => host.setState({ settingsTab: tab as any })}
                style={{
                  border: 'none', background: settingsTab === tab ? '#eff6ff' : 'transparent',
                  color: settingsTab === tab ? '#0078d4' : '#475569', fontWeight: settingsTab === tab ? 700 : 500,
                  fontSize: 13, padding: '8px 12px', borderRadius: 6, textAlign: 'left', cursor: 'pointer',
                }}
              >
                {tab}
              </button>
            ))}
          </div>

          {/* Settings Content Area */}
          <div style={{ maxWidth: 500 }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>{settingsTab}</h3>

            {settingsSavedMsg && (
              <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 13, marginBottom: 16 }}>
                ✓ Settings saved successfully.
              </div>
            )}

            {settingsTab === 'General' && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Site Title</label>
                  <input
                    type="text"
                    value={settingsForm.siteTitle}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, siteTitle: e.target.value } })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box' }}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Site Description</label>
                  <textarea
                    value={settingsForm.siteDescription}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, siteDescription: e.target.value } })}
                    style={{ width: '100%', height: 60, padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' }}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Date Format</label>
                  <select
                    value={settingsForm.dateFormat}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, dateFormat: e.target.value } })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' }}
                  >
                    <option value="MM/DD/YYYY">MM/DD/YYYY</option>
                    <option value="DD/MM/YYYY">DD/MM/YYYY</option>
                    <option value="YYYY-MM-DD">YYYY-MM-DD</option>
                  </select>
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Time Zone</label>
                  <select
                    value={settingsForm.timeZone}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, timeZone: e.target.value } })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' }}
                  >
                    <option value="(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi">(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi</option>
                    <option value="(UTC+00:00) UTC">(UTC+00:00) UTC</option>
                  </select>
                </div>

                <div style={{ marginTop: 12 }}>
                  <button
                    onClick={() => {
                      host.setState({ settingsSavedMsg: true });
                      setTimeout(() => host.setState({ settingsSavedMsg: false }), 2000);
                    }}
                    style={{
                      background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                      padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                    }}
                  >
                    Save Changes
                  </button>
                </div>
              </div>
            )}

            {settingsTab !== 'General' && (
              <div style={{ color: '#64748b', fontSize: 13 }}>
                Configuration settings for {settingsTab} are active with system default policies.
              </div>
            )}
          </div>
        </div>
      </div>
    );
}
