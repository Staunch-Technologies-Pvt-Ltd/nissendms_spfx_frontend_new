/* eslint-disable @typescript-eslint/no-unused-vars */
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
import { isMobileWidth, isTabletOrBelow } from '../responsive';

// Rich vivid accent per nav item
const NAV_ACCENTS: Record<string, string> = {
  dashboard:    '#FF2D55',  // vivid crimson-red
  list:         '#FF6B00',  // deep orange
  sites:        '#00A4EF',  // SharePoint blue
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
  sites:        '#00131A',
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
  sites:        '#4FC3F7',
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
  const viewportWidth = windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const tabletOrBelow = isTabletOrBelow(viewportWidth);
  const phone = isMobileWidth(viewportWidth);
  const desktopCollapsed = !tabletOrBelow && sidebarCollapsed;
  const mobileOpen = tabletOrBelow && !sidebarCollapsed;
  const collapsed = desktopCollapsed;
  const expandedWidth = phone ? Math.max(250, viewportWidth - 24) : 280;
  const sidebarWidth = collapsed ? 72 : expandedWidth;

  const toggleCollapsed = (): void => {
    if (tabletOrBelow) {
      host.setState({ sidebarCollapsed: mobileOpen });
      return;
    }
    host.setState({ sidebarCollapsed: !sidebarCollapsed });
  };

  const collapseBtn = (
    <button
      onClick={toggleCollapsed}
      title={collapsed ? 'Expand navigation' : 'Collapse navigation'}
      style={{
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        width: 38, height: 38, borderRadius: 10,
        border: `1.5px solid ${NAV_ACCENTS[view] || '#FF2D55'}60`,
        background: `${NAV_ACCENTS[view] || '#FF2D55'}22`,
        color: NAV_ACCENTS[view] || '#FF2D55',
        cursor: 'pointer', flexShrink: 0, transition: 'all 0.2s',
      }}
      onMouseEnter={e => { e.currentTarget.style.background = `${NAV_ACCENTS[view] || '#FF2D55'}44`; }}
      onMouseLeave={e => { e.currentTarget.style.background = `${NAV_ACCENTS[view] || '#FF2D55'}22`; }}
    >
      <Icon iconName={collapsed ? 'GlobalNavButton' : 'DoubleChevronLeft'} style={{ fontSize: 16 }} />
    </button>
  );

  const navItems: Array<{ id: AppView; label: string; iconName: string; badge?: number }> = [
    { id: 'dashboard',  label: 'Home',            iconName: 'Home' },
    { id: 'list',       label: 'Documents',        iconName: 'Documentation' },
    { id: 'sites',      label: 'Sites',             iconName: 'SharepointLogo' },
    { id: 'vessels',    label: 'Vessels',          iconName: 'Ferry' },
    { id: 'templates',  label: 'Templates & OCR',  iconName: 'FileTemplate', badge: host.state.ocrStagingCount > 0 ? host.state.ocrStagingCount : undefined },
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
          padding: collapsed ? '6px 0' : '10px 14px',
          marginBottom: 6,
          borderRadius: 14,
          boxSizing: 'border-box',
          border: active ? `1.5px solid ${accent}60` : '1.5px solid transparent',
          background: active ? `${accent}28` : 'transparent',
          color: '#fff',
          fontWeight: active ? 800 : 500,
          fontSize: isAux ? 15 : 16,
          cursor: 'pointer',
          textAlign: 'left',
          transition: 'background 0.15s ease, border-color 0.15s ease, color 0.15s ease',
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
        {active && collapsed && (
          <span style={{
            position: 'absolute', left: 1, top: '18%', bottom: '18%',
            width: 3, borderRadius: '0 3px 3px 0',
            background: accent,
            boxShadow: `0 0 8px ${accent}`,
          }} />
        )}

        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: collapsed ? 0 : 12,
          minWidth: 0,
          justifyContent: collapsed ? 'center' : 'flex-start',
          width: collapsed ? '100%' : 'auto',
        }}>
          {/* Icon container */}
          <div style={{
            position: 'relative',
            width: 42, height: 42, borderRadius: 12, flexShrink: 0,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: active ? accent : `${accent}22`,
            color: active ? '#fff' : NAV_ICON_COLORS[id as string] || accent,
            border: `2px solid ${active ? accent : accent + '50'}`,
            transition: 'all 0.2s ease',
            fontSize: isAux ? 18 : 20,
            boxShadow: active ? `0 4px 16px ${accent}99, 0 0 0 3px ${accent}22` : 'none',
          }}>
            <Icon iconName={iconName} style={{ fontSize: isAux ? 18 : 20 }} />

            {/* Badge anchored to icon container when collapsed */}
            {badgeCount !== undefined && collapsed && (
              <span style={{
                position: 'absolute',
                top: -4,
                right: -4,
                background: '#ef4444',
                color: '#fff',
                borderRadius: 10,
                minWidth: 18,
                height: 18,
                padding: '0 4px',
                fontSize: 10,
                fontWeight: 700,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                boxShadow: '0 2px 6px rgba(0,0,0,0.6)',
                border: `2px solid ${NAV_SIDEBAR_BG[view] || '#1A0A2E'}`,
                boxSizing: 'border-box',
                lineHeight: 1,
                zIndex: 2,
              }}>
                {badgeCount > 99 ? '99+' : badgeCount}
              </span>
            )}
          </div>

          <span style={{
            maxWidth: collapsed ? 0 : 180,
            opacity: collapsed ? 0 : 1,
            overflow: 'hidden',
            whiteSpace: 'nowrap',
            display: collapsed ? 'none' : 'inline-block',
            transition: 'max-width 0.25s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.15s ease',
            letterSpacing: '0.1px',
            textShadow: active ? `0 0 20px ${accent}88` : 'none',
          }}>
            {label}
          </span>
        </div>

        {/* Badge in expanded mode */}
        {badgeCount !== undefined && !collapsed && (
          <span style={{
            background: '#ef4444',
            color: '#fff',
            borderRadius: 10,
            padding: '2px 8px',
            fontSize: 11,
            fontWeight: 700,
            flexShrink: 0,
            boxShadow: '0 2px 6px rgba(239, 68, 68, 0.4)',
          }}>
            {badgeCount > 99 ? '99+' : badgeCount}
          </span>
        )}
      </button>
    );
  };

  return (
    <>
      {mobileOpen && (
        <div
          onClick={() => host.setState({ sidebarCollapsed: true })}
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(2, 6, 23, 0.45)',
            zIndex: 100000,
          }}
        />
      )}
      <div style={{
        width: sidebarWidth,
        minWidth: sidebarWidth,
        maxWidth: sidebarWidth,
        height: '100vh',
        background: NAV_SIDEBAR_BG[view] || '#1A0A2E',
        display: 'flex',
        flexDirection: 'column',
        padding: '0',
        boxSizing: 'border-box',
        userSelect: 'none',
        overflow: 'hidden',
        flexShrink: 0,
        boxShadow: `6px 0 32px rgba(0,0,0,0.6), inset -2px 0 0 ${NAV_ACCENTS[view] || '#AA00FF'}60`,
        position: tabletOrBelow ? 'fixed' : 'relative',
        top: 0,
        left: tabletOrBelow ? 0 : undefined,
        zIndex: tabletOrBelow ? 100001 : 1,
        transform: tabletOrBelow && !mobileOpen ? 'translateX(-104%)' : 'translateX(0)',
        opacity: tabletOrBelow && !mobileOpen ? 0 : 1,
        pointerEvents: tabletOrBelow && !mobileOpen ? 'none' : 'auto',
        transition: 'width 0.25s cubic-bezier(0.4, 0, 0.2, 1), min-width 0.25s cubic-bezier(0.4, 0, 0.2, 1), max-width 0.25s cubic-bezier(0.4, 0, 0.2, 1), background 0.4s ease, transform 0.25s ease, opacity 0.2s ease',
      }}>

        {/* Brand Header */}
        <div style={{
          position: 'relative', zIndex: 1,
          padding: collapsed ? '16px 0' : '18px 16px',
          borderBottom: `2px solid ${NAV_ACCENTS[view] || '#AA00FF'}50`,
          background: 'rgba(0,0,0,0.35)',
          boxShadow: `0 4px 24px ${NAV_ACCENTS[view] || '#AA00FF'}40`,
          display: 'flex', alignItems: 'center',
          gap: collapsed ? 0 : 12,
          justifyContent: collapsed ? 'center' : 'space-between',
          flexShrink: 0,
          boxSizing: 'border-box',
          width: '100%',
        }}>
          {!collapsed && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 0 }}>
              <div style={{
                width: 48, height: 48, borderRadius: 14,
                background: NAV_ACCENTS[view] || '#AA00FF',
                color: '#fff', fontWeight: 900, fontSize: 18,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                boxShadow: `0 6px 20px ${NAV_ACCENTS[view] || '#AA00FF'}BB, 0 0 0 3px ${NAV_ACCENTS[view] || '#AA00FF'}33`,
                flexShrink: 0, letterSpacing: '-0.5px',
              }}>
                VD
              </div>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 900, fontSize: 16, color: '#fff', lineHeight: 1.2, whiteSpace: 'nowrap', letterSpacing: '-0.3px' }}>
                  Vessel Documents
                </div>
                <div style={{ fontSize: 12, color: NAV_ACCENTS[view] || '#AA00FF', fontWeight: 600, whiteSpace: 'nowrap', marginTop: 2 }}>
                  Management System
                </div>
              </div>
            </div>
          )}
          {collapseBtn}
        </div>

        {/* Main Navigation */}
        <div
          className="vessel-dms-sidebar-nav"
          style={{
            position: 'relative', zIndex: 1,
            flex: 1, overflowY: 'auto', overflowX: 'hidden',
            padding: collapsed ? '12px 6px' : '12px 10px',
            boxSizing: 'border-box',
            scrollbarWidth: 'none',
            msOverflowStyle: 'none',
          }}
        >
          {/* Nav section label */}
          {!collapsed && (
            <div style={{ fontSize: 11, fontWeight: 800, color: NAV_ACCENTS[view] || '#AA00FF', textTransform: 'uppercase', letterSpacing: '2px', padding: '6px 16px 10px', marginBottom: 2 }}>
              Main Menu
            </div>
          )}
          {navItems.map(item => {
            if (item.id === 'vessels') {
              const isVesselsExpanded = !collapsed && host.state.vesselsNavExpanded;
              const isVesselsActive = view === 'vessels' || !!host.state.vesselSuggestionDialog?.open;
              const vesselsAccent = NAV_ACCENTS['vessels'] || '#F5C400';

              return (
                <div key="vessels-group" style={{ marginBottom: 6 }}>
                  <button
                    onClick={() => {
                      if (collapsed) {
                        host.setState({ sidebarCollapsed: false, vesselsNavExpanded: true });
                        host._goToView('vessels');
                      } else {
                        host.setState(prev => ({ vesselsNavExpanded: !prev.vesselsNavExpanded }));
                        if (view !== 'vessels') {
                          host._goToView('vessels');
                        }
                      }
                    }}
                    title={collapsed ? 'Vessels' : undefined}
                    style={{
                      position: 'relative',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: collapsed ? 'center' : 'space-between',
                      width: '100%',
                      padding: collapsed ? '6px 0' : '10px 14px',
                      borderRadius: 14,
                      boxSizing: 'border-box',
                      border: isVesselsActive ? `1.5px solid ${vesselsAccent}60` : '1.5px solid transparent',
                      background: isVesselsActive ? `${vesselsAccent}28` : 'transparent',
                      color: '#fff',
                      fontWeight: isVesselsActive ? 800 : 500,
                      fontSize: 16,
                      cursor: 'pointer',
                      textAlign: 'left',
                      transition: 'background 0.15s ease, border-color 0.15s ease, color 0.15s ease',
                      fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif",
                    }}
                    onMouseEnter={e => {
                      if (!isVesselsActive) {
                        e.currentTarget.style.background = `${vesselsAccent}18`;
                        e.currentTarget.style.borderColor = `${vesselsAccent}40`;
                      }
                    }}
                    onMouseLeave={e => {
                      if (!isVesselsActive) {
                        e.currentTarget.style.background = 'transparent';
                        e.currentTarget.style.borderColor = 'transparent';
                      }
                    }}
                  >
                    {/* Active left bar */}
                    {isVesselsActive && !collapsed && (
                      <span style={{
                        position: 'absolute', left: 0, top: '15%', bottom: '15%',
                        width: 4, borderRadius: '0 4px 4px 0',
                        background: vesselsAccent,
                        boxShadow: `0 0 10px ${vesselsAccent}`,
                      }} />
                    )}
                    {isVesselsActive && collapsed && (
                      <span style={{
                        position: 'absolute', left: 1, top: '18%', bottom: '18%',
                        width: 3, borderRadius: '0 3px 3px 0',
                        background: vesselsAccent,
                        boxShadow: `0 0 8px ${vesselsAccent}`,
                      }} />
                    )}

                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: collapsed ? 0 : 12,
                      minWidth: 0,
                      justifyContent: collapsed ? 'center' : 'flex-start',
                      width: collapsed ? '100%' : 'auto',
                    }}>
                      {/* Icon container */}
                      <div style={{
                        width: 42, height: 42, borderRadius: 12, flexShrink: 0,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        background: isVesselsActive ? vesselsAccent : `${vesselsAccent}22`,
                        color: isVesselsActive ? '#fff' : NAV_ICON_COLORS['vessels'] || vesselsAccent,
                        border: `2px solid ${isVesselsActive ? vesselsAccent : vesselsAccent + '50'}`,
                        transition: 'all 0.2s ease',
                        fontSize: 20,
                        boxShadow: isVesselsActive ? `0 4px 16px ${vesselsAccent}99, 0 0 0 3px ${vesselsAccent}22` : 'none',
                      }}>
                        <Icon iconName="Ferry" style={{ fontSize: 20 }} />
                      </div>

                      <span style={{
                        maxWidth: collapsed ? 0 : 180,
                        opacity: collapsed ? 0 : 1,
                        overflow: 'hidden',
                        whiteSpace: 'nowrap',
                        display: collapsed ? 'none' : 'inline-block',
                        transition: 'max-width 0.25s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.15s ease',
                        letterSpacing: '0.1px',
                        textShadow: isVesselsActive ? `0 0 20px ${vesselsAccent}88` : 'none',
                      }}>
                        Vessels
                      </span>
                    </div>

                    {!collapsed && (
                      <div style={{
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        color: isVesselsActive ? '#fff' : 'rgba(255,255,255,0.6)',
                        transform: isVesselsExpanded ? 'rotate(90deg)' : 'rotate(0deg)',
                        transition: 'transform 0.2s ease',
                        fontSize: 12,
                        width: 20, height: 20,
                      }}>
                        <Icon iconName="ChevronRight" />
                      </div>
                    )}
                  </button>

                  {/* Submenu Drilldown: Vessels Management + Suggested Vessels */}
                  {!collapsed && isVesselsExpanded && (
                    <div style={{
                      display: 'flex', flexDirection: 'column', gap: 4,
                      marginTop: 4, marginLeft: 20, paddingLeft: 14,
                      borderLeft: `2px solid ${vesselsAccent}40`,
                    }}>
                      {/* 1. Vessels Management */}
                      <button
                        onClick={() => {
                          if (host.state.vesselSuggestionDialog?.open) {
                            host.setState({ vesselSuggestionDialog: null });
                          }
                          host._goToView('vessels');
                        }}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 10,
                          padding: '9px 12px', borderRadius: 10,
                          background: (view === 'vessels' && !host.state.vesselSuggestionDialog?.open) ? `${vesselsAccent}28` : 'transparent',
                          border: (view === 'vessels' && !host.state.vesselSuggestionDialog?.open) ? `1px solid ${vesselsAccent}50` : '1px solid transparent',
                          color: (view === 'vessels' && !host.state.vesselSuggestionDialog?.open) ? '#fff' : 'rgba(255,255,255,0.8)',
                          fontWeight: (view === 'vessels' && !host.state.vesselSuggestionDialog?.open) ? 700 : 500,
                          fontSize: 13, cursor: 'pointer', textAlign: 'left',
                          transition: 'all 0.15s ease',
                          fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif",
                        }}
                        onMouseEnter={e => {
                          if (view !== 'vessels' || host.state.vesselSuggestionDialog?.open) {
                            e.currentTarget.style.background = `${vesselsAccent}15`;
                          }
                        }}
                        onMouseLeave={e => {
                          if (view !== 'vessels' || host.state.vesselSuggestionDialog?.open) {
                            e.currentTarget.style.background = 'transparent';
                          }
                        }}
                      >
                        <span style={{ fontSize: 15 }}>📋</span>
                        <span>Vessels Management</span>
                      </button>

                      {/* 2. Suggested Vessels */}
                      <button
                        onClick={() => host._openSuggestedVesselsFromSidebar()}
                        style={{
                          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                          padding: '9px 12px', borderRadius: 10,
                          background: host.state.vesselSuggestionDialog?.open ? 'linear-gradient(135deg, rgba(59,130,246,0.3), rgba(139,92,246,0.3))' : 'transparent',
                          border: host.state.vesselSuggestionDialog?.open ? '1px solid rgba(99,179,237,0.5)' : '1px solid transparent',
                          color: host.state.vesselSuggestionDialog?.open ? '#93c5fd' : 'rgba(255,255,255,0.85)',
                          fontWeight: host.state.vesselSuggestionDialog?.open ? 700 : 500,
                          fontSize: 13, cursor: 'pointer', textAlign: 'left',
                          transition: 'all 0.15s ease',
                          fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif",
                        }}
                        onMouseEnter={e => {
                          if (!host.state.vesselSuggestionDialog?.open) {
                            e.currentTarget.style.background = 'rgba(99,179,237,0.15)';
                          }
                        }}
                        onMouseLeave={e => {
                          if (!host.state.vesselSuggestionDialog?.open) {
                            e.currentTarget.style.background = 'transparent';
                          }
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                          <span style={{ fontSize: 15 }}>✨</span>
                          <span>Suggested Vessels</span>
                        </div>
                        <span style={{
                          background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)',
                          color: '#fff', fontSize: 10, fontWeight: 700,
                          padding: '1px 6px', borderRadius: 6,
                        }}>{host.state.pendingVesselSuggestions.length > 0 ? `${host.state.pendingVesselSuggestions.length}` : 'AI'}</span>
                      </button>
                    </div>
                  )}
                </div>
              );
            }
            return renderNavBtn(item.id, item.label, item.iconName, () => host._goToView(item.id), item.badge);
          })}


          {/* Aux section */}
          <div style={{
            marginTop: 14, paddingTop: 12,
            borderTop: '1px solid rgba(255,255,255,0.15)',
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
        {collapsed && !tabletOrBelow && (
          <div
            title="Vessel DMS v2.0"
            style={{
              position: 'relative', zIndex: 1,
              padding: '16px 0', borderTop: `2px solid ${NAV_ACCENTS[view] || '#AA00FF'}40`,
              background: 'rgba(0,0,0,0.30)',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >
            <span style={{
              width: 10, height: 10, borderRadius: '50%',
              background: NAV_ACCENTS[view] || '#AA00FF',
              display: 'inline-block',
              boxShadow: `0 0 10px ${NAV_ACCENTS[view] || '#AA00FF'}, 0 0 20px ${NAV_ACCENTS[view] || '#AA00FF'}88`,
            }} />
          </div>
        )}
        {!collapsed && !tabletOrBelow && (
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
    </>
  );
}
