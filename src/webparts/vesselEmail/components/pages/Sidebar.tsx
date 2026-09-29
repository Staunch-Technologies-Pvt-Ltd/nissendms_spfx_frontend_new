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
  UserItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import { isMobileWidth, isTabletOrBelow } from '../responsive';
import { clay } from '../clayTheme';

// Phase 6 — Ocean Clay: previously each module had its own vivid accent +
// matching dark sidebar background (rainbow wayfinding). Per explicit
// decision (2026-09-22), the sidebar was fully converted to the single
// Ocean Clay teal palette — these maps are kept (rather than removed) so
// every `NAV_ACCENTS[view] || fallback` call site below needs no change,
// but every module now resolves to the same teal tokens.
const NAV_ACCENTS: Record<string, string> = {
  dashboard: clay.accent, list: clay.accent, sites: clay.accent, vessels: clay.accent,
  templates: clay.accent, users: clay.accent, settings: clay.accent,
  bento_email: clay.accent, bento_compose: clay.accent, recycle: clay.accent, archive: clay.accent,
};

// Deep Harbor uses a dark rail at night and a warm ivory rail in light mode.
const NAV_SIDEBAR_BG: Record<string, string> = {
  dashboard: clay.surface, list: clay.surface, sites: clay.surface, vessels: clay.surface,
  templates: clay.surface, users: clay.surface, settings: clay.surface,
  bento_email: clay.surface, bento_compose: clay.surface, recycle: clay.surface, archive: clay.surface,
};

// Inactive icon tint — was a bright per-module hue, now a single muted clay tone.
const NAV_ICON_COLORS: Record<string, string> = {
  dashboard: clay.textMuted, list: clay.textMuted, sites: clay.textMuted, vessels: clay.textMuted,
  templates: clay.textMuted, users: clay.textMuted, settings: clay.textMuted,
  bento_email: clay.textMuted, bento_compose: clay.textMuted, recycle: clay.textMuted, archive: clay.textMuted,
};

export function renderSidebar(host: VesselEmail): React.ReactElement {
  const { view, sidebarCollapsed, windowWidth } = host.state;
  const isNight = host.state.themeMode === 'night';
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
        border: `1.5px solid ${NAV_ACCENTS[view] || clay.accent}60`,
        background: `${NAV_ACCENTS[view] || clay.accent}22`,
        color: NAV_ACCENTS[view] || clay.accent,
        cursor: 'pointer', flexShrink: 0, transition: 'all 0.2s',
      }}
      onMouseEnter={e => { e.currentTarget.style.background = `${NAV_ACCENTS[view] || clay.accent}44`; }}
      onMouseLeave={e => { e.currentTarget.style.background = `${NAV_ACCENTS[view] || clay.accent}22`; }}
    >
      <Icon iconName={collapsed ? 'GlobalNavButton' : 'DoubleChevronLeft'} style={{ fontSize: 16 }} />
    </button>
  );

  // Settings → Module Management (module_settings_api.py) can hide any of
  // these app-wide except 'settings' itself (the module catalog it serves
  // excludes both 'settings' and 'profile' on purpose, so those two entries
  // below are never filtered out — there's always a way to reopen Settings
  // and re-enable a module).
  const { hiddenModules } = host.state;
  const isHidden = (id: string): boolean => hiddenModules.indexOf(id) !== -1;

  // Typed on the literal itself (not on the .filter() result) — a type
  // annotation on the variable alone doesn't reach an array literal that's
  // immediately chained into .filter(): TS widens each `id: 'dashboard'`
  // to plain `string` before .filter() ever runs, and .filter() just
  // carries that widened `string[]` through, which no longer matches
  // `AppView`. Annotating ALL_NAV_ITEMS directly keeps every `id` as its
  // literal AppView member.
  const ALL_NAV_ITEMS: Array<{ id: AppView; label: string; iconName: string; badge?: number }> = [
    { id: 'dashboard',  label: 'Home',            iconName: 'Home' },
    { id: 'list',       label: 'Documents',        iconName: 'Documentation' },
    { id: 'sites',      label: 'Sites',             iconName: 'SharepointLogo' },
    { id: 'vessels',    label: 'Vessels',          iconName: 'Ferry' },
    // 'templates' (Templates & OCR) removed — that functionality already
    // lives directly on the SharePoint site, so this module was redundant.
    { id: 'migration',  label: 'Migration Assistant', iconName: 'MoveToFolder' },
    { id: 'users',      label: 'User Management',  iconName: 'People' },
    { id: 'settings',   label: 'Settings',         iconName: 'Settings' },
  ];
  const navItems = ALL_NAV_ITEMS.filter(item => item.id === 'settings' || !isHidden(item.id));

  // Colorful icon bg for inactive items
  const getIconBg = (id: string, active: boolean): string => {
    if (active) return NAV_ACCENTS[id] || clay.accent;
    const accent = NAV_ACCENTS[id] || clay.accent;
    return `${accent}28`; // 16% opacity tint
  };

  // Same reason as ALL_NAV_ITEMS above: annotate the literal directly, then
  // .filter() it into auxLinks, rather than annotating the .filter() result.
  const ALL_AUX_LINKS: Array<{ id: AppView | 'bento_compose'; label: string; iconName: string; onClick: () => void }> = [
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
  const auxLinks = ALL_AUX_LINKS.filter(link =>
    // 'bento_compose' isn't its own module in module_settings_api's catalog
    // (it's just "AI Bento Email" with the compose panel pre-opened), so it
    // rides on 'bento_email''s visibility instead of having its own toggle.
    link.id === 'bento_compose' ? !isHidden('bento_email') : !isHidden(link.id)
  );

  const renderNavBtn = (
    id: AppView | 'bento_compose',
    label: string,
    iconName: string,
    onClick: () => void,
    badgeCount?: number,
    isAux = false,
  ): React.ReactElement => {
    const active = view === id;
    const accent = NAV_ACCENTS[id] || clay.accent;
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
          color: active ? (isNight ? '#fff9f5' : clay.accentDark) : (isNight ? '#f4dfd0' : clay.text),
          fontWeight: active ? 800 : 500,
          fontSize: isAux ? 15 : 16,
          cursor: 'pointer',
          textAlign: 'left',
          transition: 'background 0.15s ease, border-color 0.15s ease, color 0.15s ease',
          fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif",
        }}
        onMouseEnter={e => {
          if (!active) {
            e.currentTarget.style.background = isNight ? `${accent}32` : `${accent}20`;
            e.currentTarget.style.borderColor = `${accent}40`;
            e.currentTarget.style.boxShadow = isNight ? `0 4px 14px ${accent}22` : `0 4px 14px ${accent}18`;
            e.currentTarget.style.transform = 'translateX(2px)';
          }
        }}
        onMouseLeave={e => {
          if (!active) {
            e.currentTarget.style.background = 'transparent';
            e.currentTarget.style.borderColor = 'transparent';
            e.currentTarget.style.boxShadow = 'none';
            e.currentTarget.style.transform = 'translateX(0)';
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
            textShadow: 'none',
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
        background: isNight ? '#2b211b' : (NAV_SIDEBAR_BG[view] || clay.surface),
        display: 'flex',
        flexDirection: 'column',
        padding: '0',
        boxSizing: 'border-box',
        userSelect: 'none',
        overflow: 'hidden',
        flexShrink: 0,
        boxShadow: `6px 0 24px rgba(120,190,185,0.28), inset -2px 0 0 rgba(34,134,127,0.35)`,
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
          borderBottom: `2px solid ${clay.accentSoft}`,
          background: isNight ? '#2b211b' : clay.surfaceRaised,
          boxShadow: 'none',
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
                background: NAV_ACCENTS[view] || clay.accent,
                color: '#fff', fontWeight: 900, fontSize: 18,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                boxShadow: `0 6px 20px ${NAV_ACCENTS[view] || clay.accent}BB, 0 0 0 3px ${NAV_ACCENTS[view] || clay.accent}33`,
                flexShrink: 0, letterSpacing: '-0.5px',
              }}>
                VD
              </div>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 900, fontSize: 16, color: clay.text, lineHeight: 1.2, whiteSpace: 'nowrap', letterSpacing: '-0.3px' }}>
                  Vessel Documents
                </div>
                <div style={{ fontSize: 12, color: NAV_ACCENTS[view] || clay.accent, fontWeight: 600, whiteSpace: 'nowrap', marginTop: 2 }}>
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
            <div style={{ fontSize: 11, fontWeight: 800, color: NAV_ACCENTS[view] || clay.accent, textTransform: 'uppercase', letterSpacing: '2px', padding: '6px 16px 10px', marginBottom: 2 }}>
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
                      color: isVesselsActive ? (isNight ? '#fff9f5' : clay.accentDark) : (isNight ? '#f4dfd0' : clay.text),
                      fontWeight: isVesselsActive ? 800 : 500,
                      fontSize: 16,
                      cursor: 'pointer',
                      textAlign: 'left',
                      transition: 'background 0.15s ease, border-color 0.15s ease, color 0.15s ease',
                      fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif",
                    }}
                    onMouseEnter={e => {
                      if (!isVesselsActive) {
                        e.currentTarget.style.background = isNight ? `${vesselsAccent}32` : `${vesselsAccent}20`;
                        e.currentTarget.style.borderColor = `${vesselsAccent}40`;
                        e.currentTarget.style.boxShadow = isNight ? `0 4px 14px ${vesselsAccent}22` : `0 4px 14px ${vesselsAccent}18`;
                        e.currentTarget.style.transform = 'translateX(2px)';
                      }
                    }}
                    onMouseLeave={e => {
                      if (!isVesselsActive) {
                        e.currentTarget.style.background = 'transparent';
                        e.currentTarget.style.borderColor = 'transparent';
                        e.currentTarget.style.boxShadow = 'none';
                        e.currentTarget.style.transform = 'translateX(0)';
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
                        textShadow: 'none',
                      }}>
                        Vessels
                      </span>
                    </div>

                    {!collapsed && (
                      <div style={{
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        color: isVesselsActive ? clay.accent : (isNight ? '#d8b9a1' : clay.textMuted),
                        transform: isVesselsExpanded ? 'rotate(90deg)' : 'rotate(0deg)',
                        transition: 'transform 0.2s ease',
                        fontSize: 12,
                        width: 20, height: 20,
                      }}>
                        <Icon iconName="ChevronRight" />
                      </div>
                    )}
                  </button>

                  {/* Submenu Drilldown: Vessels Management */}
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
                          color: (view === 'vessels' && !host.state.vesselSuggestionDialog?.open) ? clay.accentDark : clay.textMuted,
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
            borderTop: `1px solid ${clay.accentSoft}`,
          }}>
            {!collapsed && (
              <div style={{ fontSize: 11, fontWeight: 800, color: NAV_ACCENTS[view] || clay.accent, textTransform: 'uppercase', letterSpacing: '2px', padding: '6px 16px 10px' }}>
                Tools
              </div>
            )}
            {auxLinks.map(link =>
              renderNavBtn(link.id, link.label, link.iconName, link.onClick, undefined, true)
            )}
          </div>
        </div>

        {/* Theme switch — knob LEFT = night (dark), knob RIGHT = light */}
        {(() => {
          const isLight = host.state.themeMode !== 'night';
          const trackW = 46, trackH = 24, knob = 18, pad = 3;
          const label = isLight ? 'Light mode' : 'Night mode';
          const onKey = (e: React.KeyboardEvent<HTMLButtonElement>): void => {
            // Arrow keys follow the switch direction: left = night, right = light.
            if (e.key === 'ArrowLeft' && isLight) { e.preventDefault(); host._toggleThemeMode(); }
            if (e.key === 'ArrowRight' && !isLight) { e.preventDefault(); host._toggleThemeMode(); }
          };
          return (
            <button
              type="button"
              role="switch"
              aria-checked={isLight}
              aria-label={`Theme: ${label}. Left for night mode, right for light mode.`}
              title={isLight ? 'Switch to night mode' : 'Switch to light mode'}
              onClick={host._toggleThemeMode}
              onKeyDown={onKey}
              style={{
                position: 'relative', zIndex: 1,
                margin: collapsed ? '10px 8px' : '10px 12px',
                padding: collapsed ? '10px 0' : '8px 12px',
                minHeight: 42,
                borderRadius: 12,
                border: `1px solid ${clay.accentSoft}`,
                background: clay.surface,
                color: clay.text,
                display: 'flex', alignItems: 'center',
                justifyContent: collapsed ? 'center' : 'space-between',
                gap: 8, cursor: 'pointer', fontSize: 12, fontWeight: 700,
                transition: 'background 0.2s ease', flexShrink: 0,
              }}
            >
              {!collapsed && (
                <Icon iconName="ClearNight" style={{ fontSize: 15, color: isLight ? clay.textMuted : clay.accent }} />
              )}
              <span
                aria-hidden="true"
                style={{
                  position: 'relative', display: 'inline-block', flexShrink: 0,
                  width: trackW, height: trackH, borderRadius: trackH,
                  background: isLight ? clay.accent : '#1a120d',
                  border: `1px solid ${isLight ? clay.accentDark : '#795238'}`,
                  boxShadow: 'inset 0 1px 3px rgba(0,0,0,0.25)',
                  transition: 'background 0.25s ease, border-color 0.25s ease',
                }}
              >
                <span
                  style={{
                    position: 'absolute', top: pad - 1,
                    left: isLight ? trackW - knob - pad - 1 : pad - 1,
                    width: knob, height: knob, borderRadius: '50%',
                    background: isLight ? '#fff' : '#f8eee6',
                    boxShadow: '0 1px 3px rgba(0,0,0,0.35)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    transition: 'left 0.25s ease',
                  }}
                >
                  <Icon iconName={isLight ? 'Sunny' : 'ClearNight'} style={{ fontSize: 10, color: isLight ? clay.accentDark : '#211812' }} />
                </span>
              </span>
              {!collapsed && (
                <Icon iconName="Sunny" style={{ fontSize: 15, color: isLight ? clay.accent : clay.textMuted }} />
              )}
              {!collapsed && <span style={{ flex: 1, textAlign: 'right' }}>{label}</span>}
            </button>
          );
        })()}

        {collapsed && !tabletOrBelow && (
          <div
            title="Vessel DMS v2.0"
            style={{
              position: 'relative', zIndex: 1,
              padding: '16px 0', borderTop: `2px solid ${clay.accentSoft}`,
              background: isNight ? '#2b211b' : clay.surfaceRaised,
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >
            <span style={{
              width: 10, height: 10, borderRadius: '50%',
              background: NAV_ACCENTS[view] || clay.accent,
              display: 'inline-block',
              boxShadow: `0 0 10px ${NAV_ACCENTS[view] || clay.accent}, 0 0 20px ${NAV_ACCENTS[view] || clay.accent}88`,
            }} />
          </div>
        )}
        {!collapsed && !tabletOrBelow && (
          <div style={{
            position: 'relative', zIndex: 1,
            padding: '14px 18px', borderTop: `2px solid ${clay.accentSoft}`,
            background: isNight ? '#2b211b' : clay.surfaceRaised,
            fontSize: 12, color: isNight ? '#d8b9a1' : clay.textMuted, flexShrink: 0, fontWeight: 600,
            display: 'flex', alignItems: 'center', gap: 6,
          }}>
            <span style={{ width: 10, height: 10, borderRadius: '50%', background: NAV_ACCENTS[view] || clay.accent, display: 'inline-block', boxShadow: `0 0 10px ${NAV_ACCENTS[view] || clay.accent}, 0 0 20px ${NAV_ACCENTS[view] || clay.accent}88` }} />
            Vessel DMS v2.0
          </div>
        )}
      </div>
    </>
  );
}
