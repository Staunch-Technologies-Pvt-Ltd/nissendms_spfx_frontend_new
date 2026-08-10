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
  const { view, notificationsList, sidebarCollapsed, windowWidth } = host.state;
  const unreadNotifs = notificationsList.filter(n => !n.read).length;
  const toggleCollapsed = (): void => host.setState({ sidebarCollapsed: !sidebarCollapsed });

  const collapsed = sidebarCollapsed;
  const isNarrowScreen = (windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200)) < 640;
  const expandedWidth = isNarrowScreen ? 190 : 230;
  const sidebarWidth = collapsed ? 52 : expandedWidth;

  // Collapsed: hamburger/menu icon. Expanded: chevron (‹) to collapse.
  const collapseToggleBtn = (
    <button
      onClick={toggleCollapsed}
      title={collapsed ? 'Expand navigation' : 'Collapse navigation'}
      style={{
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        width: 28, height: 28, borderRadius: 6, border: '1px solid #e5e7eb',
        background: '#fff', color: '#6b7280', cursor: 'pointer', flexShrink: 0,
        fontSize: 13, lineHeight: 1,
      }}
    >
      {collapsed ? '☰' : String.fromCharCode(0xAB)}
    </button>
  );

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

  const auxLinks: Array<{ id: AppView | 'bento_compose'; label: string; icon: string; onClick: () => void }> = [
    { id: 'bento_email', label: 'AI Bento Email', icon: '🤖', onClick: () => host._goToView('bento_email') },
    {
      id: 'bento_compose',
      label: 'Send Email',
      icon: '✉',
      onClick: () => {
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
      },
    },
    { id: 'recycle', label: 'Recycle bin', icon: '🗑️', onClick: () => host._goToView('recycle') },
    { id: 'archive', label: 'Archive', icon: '📦', onClick: () => host._goToView('archive') },
  ];

  return (
    <div style={{
      width: sidebarWidth,
      minWidth: sidebarWidth,
      background: '#ffffff',
      borderRight: '1px solid #e0e0e0',
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'space-between',
      padding: '16px 0',
      boxSizing: 'border-box',
      userSelect: 'none',
      overflow: 'hidden',
      flexShrink: 0,
      transition: 'width 0.25s ease, min-width 0.25s ease',
    }}>
      {/* Brand Header */}
      <div>
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: collapsed ? 0 : 12,
          justifyContent: collapsed ? 'center' : 'flex-start',
          padding: collapsed ? '0 0 20px' : '0 20px 20px',
          borderBottom: '1px solid #f0f0f0',
          transition: 'padding 0.25s ease, gap 0.25s ease',
        }}>
          {!collapsed && (
            <>
              <div style={{
                width: 36, height: 36, borderRadius: 8, background: '#0078d4',
                color: '#fff', fontWeight: 700, fontSize: 16, display: 'flex',
                alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 6px rgba(0,120,212,0.3)',
                flexShrink: 0,
              }}>
                VD
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: '#111827', lineHeight: 1.2, whiteSpace: 'nowrap' }}>Vessel Documents</div>
                <div style={{ fontSize: 12, color: '#6b7280', fontWeight: 500, whiteSpace: 'nowrap' }}>Management</div>
              </div>
            </>
          )}
          {collapseToggleBtn}
        </div>

        {/* Navigation Links */}
        <div style={{ padding: collapsed ? '12px 8px' : '12px 10px', transition: 'padding 0.25s ease' }}>
          {navItems.map(item => {
            const active = view === item.id;
            return (
              <button
                key={item.id}
                onClick={() => host._goToView(item.id)}
                title={collapsed ? item.label : undefined}
                style={{
                  position: 'relative',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: collapsed ? 'center' : 'space-between',
                  width: '100%',
                  padding: collapsed ? '9px 0' : '9px 14px',
                  marginBottom: 3,
                  borderRadius: 6,
                  border: 'none',
                  background: active ? '#eff6ff' : 'transparent',
                  color: active ? '#0078d4' : '#4b5563',
                  fontWeight: active ? 600 : 500,
                  fontSize: 13,
                  cursor: 'pointer',
                  textAlign: 'left',
                  transition: 'background 0.15s ease, color 0.15s ease, padding 0.25s ease',
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: collapsed ? 0 : 10, minWidth: 0 }}>
                  <span style={{ fontSize: 16, flexShrink: 0 }}>{item.icon}</span>
                  <span style={{
                    maxWidth: collapsed ? 0 : 150,
                    opacity: collapsed ? 0 : 1,
                    overflow: 'hidden',
                    whiteSpace: 'nowrap',
                    display: 'inline-block',
                    transition: 'max-width 0.25s ease, opacity 0.15s ease',
                  }}>
                    {item.label}
                  </span>
                </div>
                {item.badge !== undefined && (
                  <span style={{
                    position: collapsed ? 'absolute' : 'static',
                    top: collapsed ? -2 : undefined,
                    right: collapsed ? -2 : undefined,
                    background: '#ef4444', color: '#fff',
                    borderRadius: collapsed ? 8 : 10,
                    padding: collapsed ? '0 4px' : '1px 7px',
                    fontSize: collapsed ? 9 : 10,
                    fontWeight: 700,
                    lineHeight: collapsed ? '14px' : undefined,
                    transition: 'all 0.2s ease',
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
      <div style={{
        borderTop: '1px solid #f0f0f0',
        padding: collapsed ? '12px 8px' : '12px 10px',
        transition: 'padding 0.25s ease',
      }}>
        {auxLinks.map(link => {
          const active = view === link.id;
          return (
            <button
              key={link.id}
              onClick={link.onClick}
              title={collapsed ? link.label : undefined}
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: collapsed ? 'center' : 'flex-start',
                gap: collapsed ? 0 : 10,
                width: '100%',
                padding: collapsed ? '9px 0' : '8px 14px',
                marginBottom: 2,
                borderRadius: 6,
                border: 'none',
                background: active ? '#eff6ff' : 'transparent',
                color: active ? '#0078d4' : '#6b7280',
                fontSize: 12,
                fontWeight: 500,
                cursor: 'pointer',
                whiteSpace: 'nowrap',
                transition: 'background 0.15s ease, color 0.15s ease, padding 0.25s ease',
              }}
            >
              <span style={{ fontSize: 15, flexShrink: 0 }}>{link.icon}</span>
              <span style={{
                maxWidth: collapsed ? 0 : 150,
                opacity: collapsed ? 0 : 1,
                overflow: 'hidden',
                whiteSpace: 'nowrap',
                display: 'inline-block',
                transition: 'max-width 0.25s ease, opacity 0.15s ease',
              }}>
                {link.label}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
