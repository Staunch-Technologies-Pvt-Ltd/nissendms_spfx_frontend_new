import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
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
  ApprovalItem, UserItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';

// Rich vivid accent per nav item
const NAV_ACCENTS: Record<string, string> = {
  dashboard:    '#FF2D55',  // vivid crimson-red
  list:         '#FF6B00',  // deep orange
  vessels:      '#F5C400',  // rich gold
  templates:    '#00C853',  // vivid green
  approvals:    '#00B0FF',  // electric blue-cyan
  reports:      '#3D5AFE',  // indigo-blue
  users:        '#AA00FF',  // vivid violet
  settings:     '#F50057',  // hot pink
  bento_email:  '#FF3D00',  // deep red-orange
  bento_compose:'#FF6D00',  // amber-orange
  recycle:      '#00BCD4',  // rich teal
  archive:      '#6200EA',  // deep purple
};

// Rich dark sidebar background per module
const NAV_SIDEBAR_BG: Record<string, string> = {
  dashboard:    '#1A0010',
  list:         '#1A0D00',
  vessels:      '#1A1500',
  templates:    '#001A08',
  approvals:    '#00101A',
  reports:      '#05082B',
  users:        '#0D0020',
  settings:     '#1A0015',
  bento_email:  '#1A0800',
  bento_compose:'#1A0E00',
  recycle:      '#001418',
  archive:      '#0A0020',
};

// Bright icon tint for inactive items
const NAV_ICON_COLORS: Record<string, string> = {
  dashboard:    '#FF6680',
  list:         '#FF9A4D',
  vessels:      '#FFD740',
  templates:    '#69F0AE',
  approvals:    '#40C4FF',
  reports:      '#7986CB',
  users:        '#CE93D8',
  settings:     '#FF80AB',
  bento_email:  '#FF6E40',
  bento_compose:'#FFAB40',
  recycle:      '#4DD0E1',
  archive:      '#B388FF',
};

export function renderSidebar(host: VesselEmail): React.ReactElement {
  const { view, sidebarCollapsed, windowWidth } = host.state;
  const toggleCollapsed = (): void => host.setState({ sidebarCollapsed: !sidebarCollapsed });

  const collapsed = sidebarCollapsed;
  const isNarrowScreen = (windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200)) < 640;
  const expandedWidth = isNarrowScreen ? 220 : 270;
  const sidebarWidth = collapsed ? 56 : expandedWidth;

  const collapseBtn = (
    <button
      onClick={toggleCollapsed}
      title={collapsed ? 'Expand navigation' : 'Collapse navigation'}
      style={{
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        width: 34, height: 34, borderRadius: 10,
        border: `1.5px solid ${NAV_ACCENTS[view] || '#FF2D55'}60`,
        background: `${NAV_ACCENTS[view] || '#FF2D55'}22`,
        color: NAV_ACCENTS[view] || '#FF2D55',
        cursor: 'pointer', flexShrink: 0, transition: 'all 0.2s',
      }}
      onMouseEnter={e => { e.currentTarget.style.background = `${NAV_ACCENTS[view] || '#FF2D55'}44`; }}
      onMouseLeave={e => { e.currentTarget.style.background = `${NAV_ACCENTS[view] || '#FF2D55'}22`; }}
    >
      <Icon iconName={collapsed ? 'GlobalNavButton' : 'DoubleChevronLeft'} style={{ fontSize: 14 }} />
    </button>
  );

  const navItems: Array<{ id: AppView; label: string; iconName: string; badge?: number }> = [
    { id: 'dashboard',  label: 'Home',            iconName: 'Home' },
    { id: 'list',       label: 'Documents',        iconName: 'Documentation' },
    { id: 'vessels',    label: 'Vessels',          iconName: 'Ferry' },
    { id: 'templates',  label: 'Templates',        iconName: 'FileTemplate' },
    { id: 'approvals',  label: 'Approvals',        iconName: 'WorkFlow' },
    { id: 'reports',    label: 'Reports',          iconName: 'BarChart4' },
    { id: 'users',      label: 'User Management',  iconName: 'People' },
    { id: 'settings',   label: 'Settings',         iconName: 'Settings' },
  ];

  // Colorful icon bg for inactive items
  const getIconBg = (id: string, active: boolean): string => {
    if (active) return NAV_ACCENTS[id] || '#0078D4';
    const accent = NAV_ACCENTS[id] || '#0078D4';
    return `${accent}28`; // 16% opacity tint
  };

  const auxLinks: Array<{ id: AppView | 'bento_compose'; label: string; iconName: string; onClick: () => void }> = [
    { id: 'bento_email',    label: 'AI Bento Email', iconName: 'Mail',       onClick: () => host._goToView('bento_email') },
    {
      id: 'bento_compose',
      label: 'Send Email',
      iconName: 'Send',
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
    { id: 'recycle', label: 'Recycle Bin', iconName: 'RecycleBin', onClick: () => host._goToView('recycle') },
    { id: 'archive', label: 'Archive',     iconName: 'Archive',    onClick: () => host._goToView('archive') },
  ];

  const renderNavBtn = (
    id: AppView | 'bento_compose',
    label: string,
    iconName: string,
    onClick: () => void,
    badgeCount?: number,
    isAux = false,
  ): React.ReactElement => {
    const active = view === id;
    const accent = NAV_ACCENTS[id] || '#0078D4';
    const iconColor = active ? '#fff' : (NAV_ICON_COLORS[id] || 'rgba(255,255,255,0.7)');

    return (
      <button
        key={id}
        onClick={onClick}
        title={collapsed ? label : undefined}
        style={{
          position: 'relative',
          display: 'flex',
          alignItems: 'center',
          justifyContent: collapsed ? 'center' : 'space-between',
          width: '100%',
          padding: collapsed ? '13px 0' : '13px 16px',
          marginBottom: 5,
          borderRadius: 14,
          border: active ? `1.5px solid ${accent}60` : '1.5px solid transparent',
          background: active ? `${accent}28` : 'transparent',
          color: '#fff',
          fontWeight: active ? 800 : 500,
          fontSize: isAux ? 15 : 16,
          cursor: 'pointer',
          textAlign: 'left',
          transition: 'background 0.15s ease, color 0.15s ease',
          fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif",
        }}
        onMouseEnter={e => {
          if (!active) {
            e.currentTarget.style.background = `${accent}18`;
            e.currentTarget.style.borderColor = `${accent}40`;
          }
        }}
        onMouseLeave={e => {
          if (!active) {
            e.currentTarget.style.background = 'transparent';
            e.currentTarget.style.borderColor = 'transparent';
          }
        }}
      >
        {/* Active left bar */}
        {active && !collapsed && (
          <span style={{
            position: 'absolute', left: 0, top: '15%', bottom: '15%',
            width: 4, borderRadius: '0 4px 4px 0',
            background: accent,
            boxShadow: `0 0 10px ${accent}`,
          }} />
        )}

        <div style={{ display: 'flex', alignItems: 'center', gap: collapsed ? 0 : 12, minWidth: 0 }}>
          {/* Icon container */}
          <div style={{
            width: 40, height: 40, borderRadius: 12, flexShrink: 0,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: active ? accent : `${accent}22`,
            color: active ? '#fff' : NAV_ICON_COLORS[id as string] || accent,
            border: `2px solid ${active ? accent : accent + '50'}`,
            transition: 'all 0.2s ease',
            fontSize: 20,
            boxShadow: active ? `0 4px 16px ${accent}99, 0 0 0 3px ${accent}22` : 'none',
          }}>
            <Icon iconName={iconName} style={{ fontSize: isAux ? 18 : 20 }} />
          </div>

          <span style={{
            maxWidth: collapsed ? 0 : 180,
            opacity: collapsed ? 0 : 1,
            overflow: 'hidden',
            whiteSpace: 'nowrap',
            display: 'inline-block',
            transition: 'max-width 0.25s ease, opacity 0.15s ease',
            letterSpacing: '0.1px',
            textShadow: active ? `0 0 20px ${accent}88` : 'none',
          }}>
            {label}
          </span>
        </div>

        {badgeCount !== undefined && (
          <span style={{
            position: collapsed ? 'absolute' : 'static',
            top: collapsed ? 2 : undefined,
            right: collapsed ? 2 : undefined,
            background: '#ef4444', color: '#fff',
            borderRadius: 10, padding: collapsed ? '0 4px' : '2px 8px',
            fontSize: collapsed ? 9 : 11, fontWeight: 700,
            lineHeight: collapsed ? '14px' : undefined,
          }}>
            {badgeCount}
          </span>
        )}
      </button>
    );
  };

  return (
    <div style={{
      width: sidebarWidth,
      minWidth: sidebarWidth,
      height: '100vh',
      background: NAV_SIDEBAR_BG[view] || '#1A0A2E',
      display: 'flex',
      flexDirection: 'column',
      padding: '0',
      boxSizing: 'border-box',
      userSelect: 'none',
      overflow: 'hidden',
      flexShrink: 0,
      transition: 'width 0.25s ease, min-width 0.25s ease, background 0.4s ease',
      boxShadow: `6px 0 32px rgba(0,0,0,0.6), inset -2px 0 0 ${NAV_ACCENTS[view] || '#AA00FF'}60`,
      position: 'relative',
    }}>

      {/* Brand Header */}
      <div style={{
        position: 'relative', zIndex: 1,
        padding: collapsed ? '20px 12px' : '20px 16px',
        borderBottom: `2px solid ${NAV_ACCENTS[view] || '#AA00FF'}50`,
        background: 'rgba(0,0,0,0.35)',
        boxShadow: `0 4px 24px ${NAV_ACCENTS[view] || '#AA00FF'}40`,
        display: 'flex', alignItems: 'center',
        gap: collapsed ? 0 : 12,
        justifyContent: collapsed ? 'center' : 'space-between',
        flexShrink: 0,
      }}>
        {!collapsed && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 0 }}>
            <div style={{
              width: 50, height: 50, borderRadius: 14,
              background: NAV_ACCENTS[view] || '#AA00FF',
              color: '#fff', fontWeight: 900, fontSize: 19,
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              boxShadow: `0 6px 20px ${NAV_ACCENTS[view] || '#AA00FF'}BB, 0 0 0 3px ${NAV_ACCENTS[view] || '#AA00FF'}33`,
              flexShrink: 0, letterSpacing: '-0.5px',
            }}>
              VD
            </div>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontWeight: 900, fontSize: 17, color: '#fff', lineHeight: 1.2, whiteSpace: 'nowrap', letterSpacing: '-0.3px' }}>
                Vessel Documents
              </div>
              <div style={{ fontSize: 13, color: NAV_ACCENTS[view] || '#AA00FF', fontWeight: 600, whiteSpace: 'nowrap', marginTop: 2 }}>
                Management System
              </div>
            </div>
          </div>
        )}
        {collapseBtn}
      </div>

      {/* Main Navigation */}
      <div style={{
        position: 'relative', zIndex: 1,
        flex: 1, overflowY: 'auto', overflowX: 'hidden',
        padding: collapsed ? '12px 8px' : '12px 10px',
      }}>
        {/* Nav section label */}
        {!collapsed && (
          <div style={{ fontSize: 11, fontWeight: 800, color: NAV_ACCENTS[view] || '#AA00FF', textTransform: 'uppercase', letterSpacing: '2px', padding: '6px 16px 10px', marginBottom: 2 }}>
            Main Menu
          </div>
        )}
        {navItems.map(item =>
          renderNavBtn(item.id, item.label, item.iconName, () => host._goToView(item.id), item.badge)
        )}

        {/* Aux section */}
        <div style={{
          marginTop: 16, paddingTop: 12,
          borderTop: '1px solid rgba(255,255,255,0.2)',
        }}>
          {!collapsed && (
            <div style={{ fontSize: 11, fontWeight: 800, color: NAV_ACCENTS[view] || '#AA00FF', textTransform: 'uppercase', letterSpacing: '2px', padding: '6px 16px 10px' }}>
              Tools
            </div>
          )}
          {auxLinks.map(link =>
            renderNavBtn(link.id, link.label, link.iconName, link.onClick, undefined, true)
          )}
        </div>
      </div>

      {/* Bottom version tag */}
      {!collapsed && (
        <div style={{
          position: 'relative', zIndex: 1,
          padding: '14px 18px', borderTop: `2px solid ${NAV_ACCENTS[view] || '#AA00FF'}40`,
          background: 'rgba(0,0,0,0.30)',
          fontSize: 12, color: 'rgba(255,255,255,0.85)', flexShrink: 0, fontWeight: 600,
          display: 'flex', alignItems: 'center', gap: 6,
        }}>
          <span style={{ width: 10, height: 10, borderRadius: '50%', background: NAV_ACCENTS[view] || '#AA00FF', display: 'inline-block', boxShadow: `0 0 10px ${NAV_ACCENTS[view] || '#AA00FF'}, 0 0 20px ${NAV_ACCENTS[view] || '#AA00FF'}88` }} />
          Vessel DMS v2.0
        </div>
      )}
    </div>
  );
}
