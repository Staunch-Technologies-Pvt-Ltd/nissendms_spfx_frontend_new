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

export function renderSidebar(host: VesselEmail): React.ReactElement {
    const { view, notificationsList } = host.state;
    const unreadNotifs = notificationsList.filter(n => !n.read).length;

    const navItems: Array<{ id: AppView; label: string; icon: string; badge?: number }> = [
      { id: 'dashboard', label: 'Home', icon: '🏠' },
      { id: 'list', label: 'Documents', icon: '📄' },
      { id: 'vessels', label: 'Vessels', icon: '🚢' },
      { id: 'templates', label: 'Templates', icon: '📑' },
      { id: 'approvals', label: 'Approvals', icon: '☑️' },
      { id: 'notifications', label: 'Notifications', icon: '🔔', badge: unreadNotifs > 0 ? unreadNotifs : undefined },
      { id: 'reports', label: 'Reports', icon: '📊' },
      { id: 'users', label: 'User Management', icon: '👥' },
      { id: 'settings', label: 'Settings', icon: '⚙️' },
    ];

    return (
      <div style={{
        width: 230,
        minWidth: 230,
        background: '#ffffff',
        borderRight: '1px solid #e0e0e0',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'space-between',
        padding: '16px 0',
        boxSizing: 'border-box',
        userSelect: 'none',
      }}>
        {/* Brand Header */}
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '0 20px 20px', borderBottom: '1px solid #f0f0f0' }}>
            <div style={{
              width: 36, height: 36, borderRadius: 8, background: '#0078d4',
              color: '#fff', fontWeight: 700, fontSize: 16, display: 'flex',
              alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 6px rgba(0,120,212,0.3)',
            }}>
              VD
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: '#111827', lineHeight: 1.2 }}>Vessel Documents</div>
              <div style={{ fontSize: 12, color: '#6b7280', fontWeight: 500 }}>Management</div>
            </div>
          </div>

          {/* Navigation Links */}
          <div style={{ padding: '12px 10px' }}>
            {navItems.map(item => {
              const active = view === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => host._goToView(item.id)}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    width: '100%',
                    padding: '9px 14px',
                    marginBottom: 3,
                    borderRadius: 6,
                    border: 'none',
                    background: active ? '#eff6ff' : 'transparent',
                    color: active ? '#0078d4' : '#4b5563',
                    fontWeight: active ? 600 : 500,
                    fontSize: 13,
                    cursor: 'pointer',
                    textAlign: 'left',
                    transition: 'all 0.15s ease',
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <span style={{ fontSize: 16 }}>{item.icon}</span>
                    <span>{item.label}</span>
                  </div>
                  {item.badge !== undefined && (
                    <span style={{
                      background: '#ef4444', color: '#fff', borderRadius: 10,
                      padding: '1px 7px', fontSize: 10, fontWeight: 700,
                    }}>
                      {item.badge}
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>

        {/* Bottom Auxiliary Links */}
        <div style={{ padding: '12px 10px', borderTop: '1px solid #f0f0f0' }}>
          <button
            onClick={() => host._goToView('bento_email')}
            style={{
              display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
              borderRadius: 6, border: 'none', background: view === 'bento_email' ? '#eff6ff' : 'transparent',
              color: view === 'bento_email' ? '#0078d4' : '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            <span>🤖</span>
            <span>AI Bento Email</span>
          </button>

          <button
            onClick={() => {
              void host._goToView('bento_email');
              host.setState({
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
              });
            }}
            style={{
              display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
              borderRadius: 6, border: 'none', background: 'transparent',
              color: '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            <span>✉</span>
            <span>Send Email</span>
          </button>

          <button
            onClick={() => host._goToView('recycle')}
            style={{
              display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
              borderRadius: 6, border: 'none', background: view === 'recycle' ? '#eff6ff' : 'transparent',
              color: view === 'recycle' ? '#0078d4' : '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            <span>🗑️</span>
            <span>Recycle bin</span>
          </button>

          <button
            onClick={() => host._goToView('archive')}
            style={{
              display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
              borderRadius: 6, border: 'none', background: view === 'archive' ? '#eff6ff' : 'transparent',
              color: view === 'archive' ? '#0078d4' : '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            <span>📦</span>
            <span>Archive</span>
          </button>
        </div>
      </div>
    );
}
