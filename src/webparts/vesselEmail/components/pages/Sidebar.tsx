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

// Enterprise Maritime: a single dark-navy rail (clay.accentDeep, so it still
// follows Settings → Color Management) with one consistent teal accent for
// the active item — no per-module colour, no gradients.
const hexRgb = (hex: string): number[] => {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};
const rgbaOf = (hex: string, a: number): string => `rgba(${hexRgb(hex).join(',')},${a})`;
const SB_FONT_UI = "'Manrope', 'Segoe UI Variable', 'Segoe UI', sans-serif";
const SB_FONT_DISPLAY = "'Sora', 'Segoe UI Variable', 'Segoe UI', sans-serif";
const SB_FONT_MONO = "'JetBrains Mono', Consolas, monospace";
// Hex fallback mirroring clayLight.accent — used only where a literal hex is
// required (rgba() math); `clay.accent` (the CSS-var form) is used everywhere
// else so Settings → Color Management keeps re-theming this correctly.
const ACCENT_HEX = '#0e7490';
// Sidebar text is always light — the rail background (clay.accentDeep) is
// always dark, in both light and night app themes.
const SB_TEXT = '#f3f7fb';
const SB_TEXT_MUTED = 'rgba(243,247,251,0.6)';
const SB_DIVIDER = 'rgba(243,247,251,0.14)';

export function renderSidebar(host: VesselEmail): React.ReactElement {
  const { view, sidebarCollapsed, windowWidth } = host.state;
  const isNight = host.state.themeMode === 'night';
  const viewportWidth = windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const tabletOrBelow = isTabletOrBelow(viewportWidth);
  const phone = isMobileWidth(viewportWidth);
  const desktopCollapsed = !tabletOrBelow && sidebarCollapsed;
  const mobileOpen = tabletOrBelow && !sidebarCollapsed;
  const collapsed = desktopCollapsed;
  const expandedWidth = phone ? Math.max(250, viewportWidth - 24) : 320;
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
        width: 34, height: 34, borderRadius: 8,
        border: '1px solid rgba(255,255,255,0.18)',
        background: 'rgba(255,255,255,0.08)',
        color: '#ffffff',
        cursor: 'pointer', flexShrink: 0, transition: 'background 0.2s ease',
      }}
      onMouseEnter={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.16)'; }}
      onMouseLeave={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.08)'; }}
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
    { id: 'dashboard',  label: 'Dashboard',           iconName: 'Home' },
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

  // Same reason as ALL_NAV_ITEMS above: annotate the literal directly, then
  // .filter() it into auxLinks, rather than annotating the .filter() result.
  const ALL_AUX_LINKS: Array<{ id: AppView | 'bento_compose'; label: string; iconName: string; onClick: () => void }> = [
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
          minHeight: 50,
          padding: collapsed ? '6px 0' : '5px 12px 5px 16px',
          marginBottom: 6,
          borderRadius: 8,
          boxSizing: 'border-box',
          border: '1px solid transparent',
          background: active ? rgbaOf(ACCENT_HEX, 0.22) : 'transparent',
          color: active ? '#ffffff' : SB_TEXT,
          fontWeight: active ? 700 : 500,
          fontSize: 14.5,
          cursor: 'pointer',
          textAlign: 'left',
          transition: 'background 150ms ease, color 150ms ease',
          fontFamily: SB_FONT_UI,
        }}
        onMouseEnter={e => {
          if (!active) {
            e.currentTarget.style.background = 'rgba(255,255,255,0.07)';
          }
        }}
        onMouseLeave={e => {
          if (!active) {
            e.currentTarget.style.background = 'transparent';
          }
        }}
      >
        {/* Thin left accent indicator — replaces a solid bright fill for the active row */}
        {active && !collapsed && (
          <span aria-hidden="true" style={{ position: 'absolute', left: 0, top: 6, bottom: 6, width: 3, borderRadius: '0 3px 3px 0', background: clay.accent }} />
        )}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: collapsed ? 0 : 11,
          minWidth: 0,
          justifyContent: collapsed ? 'center' : 'flex-start',
          width: collapsed ? '100%' : 'auto',
        }}>
          {/* Icon container — flat tile, single accent, no per-module colour */}
          <div style={{
            position: 'relative',
            width: 32, height: 32, borderRadius: 8, flexShrink: 0,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: active ? 'rgba(255,255,255,0.16)' : 'rgba(255,255,255,0.07)',
            color: '#ffffff',
            transition: 'background 150ms ease',
          }}>
            <Icon iconName={iconName} style={{ fontSize: 17 }} />

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
                border: `2px solid ${clay.accentDeep}`,
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
            transition: 'max-width 0.2s ease, opacity 0.15s ease',
            letterSpacing: '0.1px',
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
      <div className="vessel-dms-sidebar" style={{
        width: sidebarWidth,
        minWidth: sidebarWidth,
        maxWidth: sidebarWidth,
        height: '100vh',
        maxHeight: '100dvh',
        background: clay.accentDeep,
        borderRight: '1px solid rgba(255,255,255,0.08)',
        boxShadow: tabletOrBelow ? '10px 0 28px rgba(2,10,24,0.14)' : 'none',
        display: 'flex',
        flexDirection: 'column',
        padding: '0',
        boxSizing: 'border-box',
        userSelect: 'none',
        overflow: 'hidden',
        flexShrink: 0,
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
          padding: collapsed ? '18px 0' : '20px 18px',
          borderBottom: `1px solid ${SB_DIVIDER}`,
          display: 'flex', alignItems: 'center',
          gap: collapsed ? 0 : 10,
          justifyContent: collapsed ? 'center' : 'space-between',
          flexShrink: 0,
          boxSizing: 'border-box',
          width: '100%',
          minHeight: 72,
        }}>
          {!collapsed && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 0 }}>
              {/* Static brand mark — Vessel DMS */}
              <div
                aria-label="Vessel DMS"
                role="img"
                style={{
                  width: 38, height: 38, borderRadius: 9, flexShrink: 0,
                  background: rgbaOf(ACCENT_HEX, 0.28),
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  color: '#ffffff',
                }}
              >
                <Icon iconName="Ferry" style={{ fontSize: 18 }} />
              </div>
              <div style={{ minWidth: 0, display: 'flex', alignItems: 'center' }}>
                <div style={{ fontFamily: SB_FONT_DISPLAY, fontWeight: 700, fontSize: 18, color: SB_TEXT, lineHeight: 1.2, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', letterSpacing: '-0.015em' }}>
                  Vessel DMS
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
            padding: collapsed ? '16px 8px' : '18px 14px',
            boxSizing: 'border-box',
            scrollbarWidth: 'none',
            msOverflowStyle: 'none',
          }}
        >
          {/* Nav section label */}
          {!collapsed && (
            <div style={{ fontFamily: SB_FONT_MONO, fontSize: 10, fontWeight: 700, color: 'rgba(243,247,251,0.38)', textTransform: 'uppercase', letterSpacing: '0.12em', padding: '2px 4px 9px', marginBottom: 1 }}>
              Main
            </div>
          )}
          {navItems.map(item => {
            if (item.id === 'vessels') {
              const isVesselsExpanded = !collapsed && host.state.vesselsNavExpanded;
              const isVesselsActive = view === 'vessels' || !!host.state.vesselSuggestionDialog?.open;

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
                      minHeight: 50,
                      padding: collapsed ? '6px 0' : '5px 12px 5px 16px',
                      borderRadius: 8,
                      boxSizing: 'border-box',
                      border: '1px solid transparent',
                      background: isVesselsActive ? rgbaOf(ACCENT_HEX, 0.22) : 'transparent',
                      color: isVesselsActive ? '#ffffff' : SB_TEXT,
                      fontWeight: isVesselsActive ? 700 : 500,
                      fontSize: 14.5,
                      cursor: 'pointer',
                      textAlign: 'left',
                      transition: 'background 150ms ease, color 150ms ease',
                      fontFamily: SB_FONT_UI,
                    }}
                    onMouseEnter={e => {
                      if (!isVesselsActive) {
                        e.currentTarget.style.background = 'rgba(255,255,255,0.07)';
                      }
                    }}
                    onMouseLeave={e => {
                      if (!isVesselsActive) {
                        e.currentTarget.style.background = 'transparent';
                      }
                    }}
                  >
                    {isVesselsActive && !collapsed && (
                      <span aria-hidden="true" style={{ position: 'absolute', left: 0, top: 6, bottom: 6, width: 3, borderRadius: '0 3px 3px 0', background: clay.accent }} />
                    )}
                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: collapsed ? 0 : 11,
                      minWidth: 0,
                      justifyContent: collapsed ? 'center' : 'flex-start',
                      width: collapsed ? '100%' : 'auto',
                    }}>
                      {/* Icon container — flat tile, matches the rest of the nav */}
                      <div style={{
                        width: 32, height: 32, borderRadius: 8, flexShrink: 0,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        background: isVesselsActive ? 'rgba(255,255,255,0.16)' : 'rgba(255,255,255,0.07)',
                        color: '#ffffff',
                        transition: 'background 150ms ease',
                      }}>
                        <Icon iconName="Ferry" style={{ fontSize: 17 }} />
                      </div>

                      <span style={{
                        maxWidth: collapsed ? 0 : 180,
                        opacity: collapsed ? 0 : 1,
                        overflow: 'hidden',
                        whiteSpace: 'nowrap',
                        display: collapsed ? 'none' : 'inline-block',
                        transition: 'max-width 0.2s ease, opacity 0.15s ease',
                        letterSpacing: '0.1px',
                      }}>
                        Vessels
                      </span>
                    </div>

                    {!collapsed && (
                      <div style={{
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        color: isVesselsActive ? '#ffffff' : SB_TEXT_MUTED,
                        transform: isVesselsExpanded ? 'rotate(90deg)' : 'rotate(0deg)',
                        transition: 'transform 0.2s ease',
                        fontSize: 12,
                        width: 20, height: 20,
                      }}>
                        <Icon iconName="ChevronRight" />
                      </div>
                    )}
                  </button>

                  {/* Submenu Drilldown: Vessels Management — always mounted while the
                      sidebar is expanded, animated via max-height/opacity so it slides
                      shut smoothly instead of popping in/out. */}
                  {!collapsed && (
                    <div style={{
                      display: 'flex', flexDirection: 'column', gap: 2,
                      marginLeft: 30, paddingLeft: 12,
                      borderLeft: `1px solid ${SB_DIVIDER}`,
                      maxHeight: isVesselsExpanded ? 44 : 0,
                      opacity: isVesselsExpanded ? 1 : 0,
                      marginTop: isVesselsExpanded ? 4 : 0,
                      overflow: 'hidden',
                      transition: 'max-height 200ms ease, opacity 150ms ease, margin-top 200ms ease',
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
                          position: 'relative',
                          display: 'flex', alignItems: 'center', gap: 9,
                          padding: '8px 10px', borderRadius: 7,
                          background: view === 'vessels' ? rgbaOf(ACCENT_HEX, 0.18) : 'transparent',
                          border: 'none',
                          color: view === 'vessels' ? '#ffffff' : SB_TEXT_MUTED,
                          fontWeight: view === 'vessels' ? 700 : 600,
                          fontSize: 13, cursor: 'pointer', textAlign: 'left',
                          transition: 'background 150ms ease, color 150ms ease',
                          fontFamily: SB_FONT_UI,
                        }}
                        onMouseEnter={e => { if (view !== 'vessels') e.currentTarget.style.background = 'rgba(255,255,255,0.07)'; }}
                        onMouseLeave={e => { if (view !== 'vessels') e.currentTarget.style.background = 'transparent'; }}
                      >
                        <Icon iconName="BulletedList" style={{ fontSize: 13 }} />
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
            marginTop: 12, paddingTop: 12,
            borderTop: `1px solid ${SB_DIVIDER}`,
          }}>
            {!collapsed && (
              <div style={{ fontFamily: SB_FONT_MONO, fontSize: 10, fontWeight: 700, color: 'rgba(243,247,251,0.38)', textTransform: 'uppercase', letterSpacing: '0.12em', padding: '2px 4px 9px' }}>
                Tools
              </div>
            )}
            {auxLinks.map(link =>
              renderNavBtn(link.id, link.label, link.iconName, link.onClick, undefined, true)
            )}
          </div>
        </div>

        {collapsed && !tabletOrBelow && (
          <div
            title="Vessel DMS v2.0"
            style={{
              position: 'relative', zIndex: 1,
              padding: '11px 0', borderTop: `1px solid ${SB_DIVIDER}`,
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >
            <span style={{ width: 7, height: 7, borderRadius: '50%', background: '#2dd4bf', display: 'inline-block' }} />
          </div>
        )}
        {!collapsed && !tabletOrBelow && (
          <div style={{
            position: 'relative', zIndex: 1,
            padding: '9px 16px', borderTop: `1px solid ${SB_DIVIDER}`,
            fontFamily: SB_FONT_MONO, fontSize: 10.5, color: 'rgba(243,247,251,0.45)', flexShrink: 0, fontWeight: 600,
            display: 'flex', alignItems: 'center', gap: 7,
          }}>
            <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#2dd4bf', display: 'inline-block', flexShrink: 0 }} />
            Vessel DMS v2.0
          </div>
        )}
      </div>
    </>
  );
}
