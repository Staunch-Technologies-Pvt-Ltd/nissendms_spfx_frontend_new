import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import {
  badge, GROUP_COLORS, DATASOURCE_TAGS_MAP, VESSEL_TYPES, cleanName, suggestTagFromFilename,
  INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES,
  INITIAL_MOCK_NOTIFICATIONS, INITIAL_MOCK_USERS,
} from '../constants';
import type {
  FlatRow, GroupedRow, VesselRecord,
} from '../types/rows';
import type { BentoEmailLog } from '../types/bento';
import type { AppView, ModalMode } from '../types/view';
import type {
  FormState, DocPreviewItem, DeletedNode, DocumentItem, TemplateItem,
  NotificationItem, UserItem, AlertItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import { isMobileWidth, isTabletOrBelow } from '../responsive';
import { ThemeProvider } from '@fluentui/react';
import { clay, buildDeepHarborTheme } from '../clayTheme';

// Phase 6 — Ocean Clay: previously each module had its own vivid top-bar
// accent (matching Sidebar.tsx's rainbow NAV_ACCENTS). Per explicit decision
// (2026-09-22), fully converted to the single Ocean Clay teal — map kept so
// every VIEW_ACCENTS[view] call site below needs no change.
const VIEW_ACCENTS: Record<string, string> = {
  dashboard: clay.accent, list: clay.accent, vessels: clay.accent, templates: clay.accent,
  users: clay.accent, settings: clay.accent, bento_email: clay.accent, email_notify: clay.accent,
  bento_compose: clay.accent, recycle: clay.accent, archive: clay.accent, migration: clay.accent,
};

// View → display label
const VIEW_LABELS: Record<string, string> = {
  dashboard: 'Dashboard',
  list: 'Documents',
  vessels: 'Vessels',
  templates: 'Templates',
  migration: 'Migration Assistant',
  users: 'User Management',
  settings: 'Settings',
  profile: 'Profile',
  bento_email: 'AI Bento Email',
  email_notify: 'AI Bento Email',
  recycle: 'Recycle Bin',
  archive: 'Archive',
};

// Inject global CSS once to break out of SharePoint workbench constraints
function injectFullScreenStyles(): void {
  const id = 'vessel-dms-fullscreen';
  let style = document.getElementById(id) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement('style');
    style.id = id;
    document.head.appendChild(style);
  }
  style.textContent = `
    /* ── Whole-app dark mode tokens ──────────────────────────────────────
       Every page still sets its own inline colors (no shared component
       library to retheme centrally — see clayTheme.ts header comment), so
       the bulk of those inline hex literals were swapped for var(--vdms-*)
       references instead of being duplicated with isNight ternaries. These
       custom properties are the single place that actually flips between
       light and night, scoped to the [data-vessel-theme] attribute set on
       the app root below. Night values reuse the existing warm-brown
       "Ocean Clay" night palette (see deepHarborNightTheme / the night
       input styles further down) so nothing clashes. */
    [data-vessel-theme="light"] {
      --vdms-surface: #ffffff;
      --vdms-surface-alt: #f8fafc;
      --vdms-border: #e2e8f0;
      --vdms-border-soft: #f1f5f9;
      --vdms-text: #0f172a;
      --vdms-text-secondary: #475569;
      --vdms-text-muted: #64748b;
      --vdms-text-faint: #94a3b8;
      --vdms-toggle-active-bg: #0f172a;
      --vdms-toggle-active-text: #ffffff;
      /* clay.* tokens (clayTheme.ts) */
      --clay-bg: #f8f1ea;
      --clay-surface: #fff9f5;
      --clay-surface-raised: #f7e6d8;
      --clay-surface-hover: #f2e2d5;
      --clay-text: #342417;
      --clay-text-muted: #8a6552;
      --clay-accent-soft: #f7d8bf;
      --clay-accent-soft-hover: #f0cab1;
      --clay-icon-bg: linear-gradient(150deg, #f9dcc0, #efb57a);
      --clay-shadow-raised: 8px 8px 18px rgba(221,145,89,0.22), -8px -8px 16px rgba(255,255,255,0.88);
      --clay-shadow-raised-hover: 10px 10px 22px rgba(221,145,89,0.28), -10px -10px 20px rgba(255,255,255,0.92);
      --clay-shadow-button: 0 10px 22px rgba(221,145,89,0.35), inset 0 2px 3px rgba(255,255,255,0.45), inset 0 -3px 6px rgba(150,89,42,0.28);
      --clay-shadow-icon: inset 0 2px 3px rgba(255,255,255,0.7), inset 0 -3px 5px rgba(185,110,53,0.22);
      --clay-pill-active-bg: #cdeedb;
      --clay-pill-active-text: #245a3d;
      --clay-pill-active-shadow: inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -2px 3px rgba(36,90,61,0.18);
      --clay-pill-warn-bg: #e3d9c2;
      --clay-pill-warn-text: #7a6420;
      --clay-pill-warn-shadow: inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -2px 3px rgba(122,100,32,0.18);
      --clay-pill-danger-bg: #ecccc8;
      --clay-pill-danger-text: #8a3226;
      --clay-pill-danger-shadow: inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -2px 3px rgba(138,50,38,0.18);
    }
    [data-vessel-theme="night"] {
      --vdms-surface: #2b211b;
      --vdms-surface-alt: #3a291f;
      --vdms-border: #493225;
      --vdms-border-soft: #3a291f;
      --vdms-text: #f8eee6;
      --vdms-text-secondary: #d8c4b3;
      --vdms-text-muted: #c7a58d;
      --vdms-text-faint: #a8886f;
      --vdms-toggle-active-bg: #DD9159;
      --vdms-toggle-active-text: #211812;
      /* clay.* tokens (clayTheme.ts) — night values */
      --clay-bg: #211812;
      --clay-surface: #2b211b;
      --clay-surface-raised: #3a291f;
      --clay-surface-hover: #45311f;
      --clay-text: #f8eee6;
      --clay-text-muted: #c7a58d;
      --clay-accent-soft: #5a3a24;
      --clay-accent-soft-hover: #6a4429;
      --clay-icon-bg: linear-gradient(150deg, #5a3a24, #7a4a2a);
      --clay-shadow-raised: 6px 6px 14px rgba(0,0,0,0.45), -4px -4px 10px rgba(255,255,255,0.03);
      --clay-shadow-raised-hover: 8px 8px 18px rgba(0,0,0,0.55), -5px -5px 12px rgba(255,255,255,0.04);
      --clay-shadow-button: 0 8px 18px rgba(0,0,0,0.45), inset 0 1px 2px rgba(255,255,255,0.18), inset 0 -3px 6px rgba(0,0,0,0.3);
      --clay-shadow-icon: inset 0 1px 2px rgba(255,255,255,0.12), inset 0 -3px 5px rgba(0,0,0,0.35);
      --clay-pill-active-bg: #1f3b2c;
      --clay-pill-active-text: #9fe0bb;
      --clay-pill-active-shadow: inset 0 1px 2px rgba(255,255,255,0.06), inset 0 -2px 3px rgba(0,0,0,0.3);
      --clay-pill-warn-bg: #3d3420;
      --clay-pill-warn-text: #e5c97a;
      --clay-pill-warn-shadow: inset 0 1px 2px rgba(255,255,255,0.06), inset 0 -2px 3px rgba(0,0,0,0.3);
      --clay-pill-danger-bg: #45231f;
      --clay-pill-danger-text: #f2a79c;
      --clay-pill-danger-shadow: inset 0 1px 2px rgba(255,255,255,0.06), inset 0 -2px 3px rgba(0,0,0,0.3);
    }

    /* Force web part zone to not clip our fixed overlay */
    .CanvasZone, .CanvasSection, .CanvasComponent,
    [data-automation-id="CanvasControl"],
    .ms-SPCanvas, .SPCanvas, .workbenchPageContent,
    .sp-Canvas, .sp-CanvasSection {
      overflow: visible !important;
    }
    /* Ensure html/body don't clip */
    html, body {
      overflow: hidden !important;
    }
    /* Sidebar nav scrollbar hiding */
    .vessel-dms-sidebar-nav::-webkit-scrollbar {
      display: none;
      width: 0px;
      height: 0px;
    }
    .vessel-dms-sidebar-nav {
      -ms-overflow-style: none;
      scrollbar-width: none;
    }

    /* Night mode keeps the warm clay accent while lowering the surrounding surfaces. */
    [data-vessel-theme="night"] input,
    [data-vessel-theme="night"] select,
    [data-vessel-theme="night"] textarea {
      background: #2b211b !important;
      color: #f8eee6 !important;
      border-color: #795238 !important;
      color-scheme: dark;
    }
    [data-vessel-theme="night"] input::placeholder,
    [data-vessel-theme="night"] textarea::placeholder {
      color: #b99a84 !important;
    }

    /* Shared interaction feedback for buttons and custom clickable surfaces. */
    .vessel-dms-app button,
    .vessel-dms-app a,
    .vessel-dms-app [role="button"] {
      transition: filter 0.16s ease, transform 0.16s ease, box-shadow 0.16s ease;
    }
    .vessel-dms-app button:hover:not(:disabled),
    .vessel-dms-app a:hover,
    .vessel-dms-app [role="button"]:hover {
      filter: brightness(0.94) saturate(1.08) !important;
      transform: translateY(-1px) !important;
      outline: 2px solid rgba(221,145,89,0.72) !important;
      outline-offset: 2px;
      box-shadow: 0 0 0 4px rgba(221,145,89,0.18) !important;
    }
    .vessel-dms-app button:active:not(:disabled),
    .vessel-dms-app a:active,
    .vessel-dms-app [role="button"]:active {
      filter: brightness(0.88) saturate(1.12);
      transform: translateY(0);
    }
    .vessel-dms-app button:focus-visible,
    .vessel-dms-app a:focus-visible,
    .vessel-dms-app [role="button"]:focus-visible {
      outline: 3px solid rgba(221,145,89,0.55);
      outline-offset: 2px;
    }
  `;
}

/** Live deletion popup (Part 3): a transient toast for any newly-seen
 * vessel_deleted / document_deleted alert — the same deletion_log-backed
 * rows the alert bell already lists (GET /api/alerts/all, polled every
 * 30s by host._fetchAlerts). Shows who deleted it, what, and which site,
 * auto-dismisses after 10s, can be dismissed early, and clicking it
 * deep-links into the Recycle Bin — reusing the exact same handling the
 * bell dropdown's own "View in Recycle Bin" buttons already use. */
function DeletionToastLayer({ host }: { host: VesselEmail }): React.ReactElement | null {
  const alertsList = host.state.alertsList || [];
  const alertsLoaded = host.state.alertsLoaded;
  const seenIds = React.useRef<Set<string>>(new Set());
  const initialized = React.useRef(false);
  const [toasts, setToasts] = React.useState<AlertItem[]>([]);

  React.useEffect(() => {
    const deletions = alertsList.filter(
      a => a.alert_type === 'vessel_deleted' || a.alert_type === 'document_deleted'
    );
    // GET /api/alerts/all resolves after this component's first render, so
    // alertsList is still [] on mount. Gating "first load" on alertsLoaded
    // (not just "have we run once") means we keep re-baselining seenIds —
    // without toasting anything — through every render before the real
    // data arrives, instead of seeding an empty baseline too early and then
    // treating every pre-existing deletion as brand new the moment the
    // actual list shows up (this is what fired a toast for every old
    // deletion in the log on page open).
    if (!alertsLoaded || !initialized.current) {
      deletions.forEach(a => seenIds.current.add(a.id));
      if (alertsLoaded) initialized.current = true;
      return;
    }
    const fresh = deletions.filter(a => !seenIds.current.has(a.id));
    if (!fresh.length) return;
    fresh.forEach(a => seenIds.current.add(a.id));
    setToasts(prev => [...fresh, ...prev].slice(0, 3));
    fresh.forEach(a => {
      window.setTimeout(() => {
        setToasts(prev => prev.filter(t => t.id !== a.id));
      }, 10000);
    });
  }, [alertsList, alertsLoaded]);

  if (!toasts.length) return null;

  return (
    <div style={{ position: 'fixed', top: 90, right: 20, zIndex: 200000, display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 340 }}>
      {toasts.map(t => (
        <div key={t.id} style={{ background: '#fff', border: '1px solid #fecdd3', borderRadius: 12, boxShadow: '0 12px 32px rgba(15,23,42,0.25)', padding: '14px 16px', display: 'flex', gap: 10 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: '#fee2e2', color: '#991b1b', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 17, flexShrink: 0 }}>
            <Icon iconName="Delete" />
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: '#0f172a' }}>
              {t.alert_type === 'vessel_deleted' ? 'Vessel deleted' : 'Item deleted'}
            </div>
            <div style={{ fontSize: 12, color: '#475569', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={t.folder_name}>
              {t.folder_name}
            </div>
            <div style={{ fontSize: 11, color: '#64748b', marginTop: 3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {t.created_by_name || t.created_by_email || 'Unknown user'}
              {t.site_name ? ` • ${t.site_name}` : ''}
            </div>
            <button
              onClick={() => {
                if (!t.read) host._markAlertRead(t.id);
                setToasts(prev => prev.filter(x => x.id !== t.id));
                void host._goToView('recycle').catch(() => undefined);
              }}
              style={{ marginTop: 8, padding: '4px 10px', borderRadius: 6, border: '1px solid #fca5a5', background: '#fff5f5', color: '#dc2626', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}
            >
              View in Recycle Bin
            </button>
          </div>
          <button
            onClick={() => setToasts(prev => prev.filter(x => x.id !== t.id))}
            aria-label="Dismiss"
            style={{ border: 'none', background: 'transparent', color: '#94a3b8', cursor: 'pointer', fontSize: 14, padding: 0, alignSelf: 'flex-start' }}
          >
            ✕
          </button>
        </div>
      ))}
    </div>
  );
}

export function renderLayout(host: VesselEmail, content: React.ReactElement): React.ReactElement {
  const userDisplayName = host.props.userDisplayName || 'Admin';
  const isNight = host.state.themeMode === 'night';
  const isWorkspaceFullScreen = host.state.fullScreenWorkspace && host.state.view === 'list';
  const viewLabel = VIEW_LABELS[host.state.view] || host.state.view.replace(/_/g, ' ');
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const tabletOrBelow = isTabletOrBelow(viewportWidth);
  const phone = isMobileWidth(viewportWidth);

  injectFullScreenStyles();

  // Settings → Color Management (color_settings_api.py) recolors the app at
  // runtime; the clay.* tokens above already pick that up via CSS custom
  // properties (applyColorTheme), and Fluent's own theme is rebuilt here
  // from the same saved colors so its handful of stock controls match.
  const colorSet = isNight ? host.state.colorTheme.night : host.state.colorTheme.light;
  const fluentTheme = buildDeepHarborTheme(colorSet, isNight);

  return (
    <ThemeProvider theme={fluentTheme}>
    <div className="vessel-dms-app" style={{
      display: 'flex', alignItems: 'stretch',
      height: '100vh', minHeight: '100vh',
      width: '100vw', overflow: 'hidden',
      background: clay.bg,
      fontFamily: "'Segoe UI Variable', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 99999,
    }} data-vessel-theme={host.state.themeMode}>
      <DeletionToastLayer host={host} />
      {!isWorkspaceFullScreen && host._renderSidebar()}

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0, overflow: 'hidden' }}>

        {/* ── Top Bar ── */}
        <div style={{
          height: 76, flexShrink: 0,
          background: isNight ? clay.accentDeep : (VIEW_ACCENTS[host.state.view] || clay.accent),
          boxShadow: `0 4px 24px ${clay.accentGlow}`,
          transition: 'background 0.4s ease, box-shadow 0.4s ease',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          padding: phone ? '0 12px' : tabletOrBelow ? '0 18px' : '0 36px',
          gap: 10,
        }}>
          {/* Breadcrumb */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            {tabletOrBelow && (
              <button
                type="button"
                onClick={() => host.setState({ sidebarCollapsed: false })}
                style={{
                  width: 42,
                  height: 42,
                  borderRadius: 10,
                  border: '1px solid rgba(255,255,255,0.5)',
                  background: 'rgba(255,255,255,0.15)',
                  color: '#fff',
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  cursor: 'pointer',
                  flexShrink: 0,
                  transition: 'all 0.15s ease',
                }}
                onMouseEnter={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.26)'; }}
                onMouseLeave={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.15)'; }}
                aria-label="Open navigation"
              >
                <Icon iconName="GlobalNavButton" style={{ fontSize: 18 }} />
              </button>
            )}
            {!phone && <span style={{ color: 'rgba(255,255,255,0.75)', fontSize: 16, fontWeight: 600, letterSpacing: '0.2px' }}>Vessel DMS</span>}
            <Icon iconName="ChevronRight" style={{ fontSize: 14, color: 'rgba(255,255,255,0.55)' }} />
            <span style={{ color: '#ffffff', fontSize: phone ? 18 : 24, fontWeight: 900, letterSpacing: '-0.5px', textShadow: '0 2px 12px rgba(0,0,0,0.25)' }}>
              {viewLabel}
            </span>
            <button
              type="button"
              onClick={() => { void host._refreshCurrentModule(); }}
              title={`Refresh ${viewLabel}`}
              aria-label={`Refresh ${viewLabel}`}
              style={{
                width: 34, height: 34, marginLeft: 4, borderRadius: '50%',
                border: '1px solid rgba(255,255,255,0.62)',
                background: 'rgba(255,255,255,0.2)', color: '#fff',
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                cursor: 'pointer', flexShrink: 0,
              }}
            >
              <Icon iconName="Refresh" style={{ fontSize: 15 }} />
            </button>
          </div>

          {/* Right controls */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {host.state.view === 'list' && (
              <button
                type="button"
                onClick={host._toggleFullScreenWorkspace}
                title={isWorkspaceFullScreen ? 'Exit full-screen workspace' : 'Open full-screen workspace'}
                aria-label={isWorkspaceFullScreen ? 'Exit full-screen workspace' : 'Open full-screen workspace'}
                style={{
                  width: 42, height: 42, borderRadius: 12,
                  border: '1px solid rgba(255,255,255,0.4)',
                  background: isWorkspaceFullScreen ? 'rgba(255,255,255,0.3)' : 'rgba(255,255,255,0.15)',
                  color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                  cursor: 'pointer', boxShadow: '0 3px 10px rgba(0,0,0,0.15)',
                }}
              >
                <Icon iconName={isWorkspaceFullScreen ? 'BackToWindow' : 'FullScreen'} style={{ fontSize: 17 }} />
              </button>
            )}
            {/* Top-bar Search box removed per request — it was dead UI on
                every module (no value/onChange ever wired to it; see git
                history), and each module already has its own real search
                (the "Search vessel, file name..." box in Documents, etc.). */}

            {/* ── Alert Bell ── */}
            <div style={{ position: 'relative' }}>
              <button
                type="button"
                data-alert-bell="true"
                onClick={e => { e.stopPropagation(); host._toggleAlertBell(); }}
                style={{
                  width: 44, height: 44, borderRadius: '50%', border: 'none',
                  background: host.state.alertOpen ? 'rgba(255,255,255,0.35)' : 'rgba(255,255,255,0.18)',
                  color: '#fff',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  cursor: 'pointer', position: 'relative',
                  transition: 'background 0.15s ease',
                  boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
                }}
                onMouseEnter={e => { if (!host.state.alertOpen) e.currentTarget.style.background = 'rgba(255,255,255,0.25)'; }}
                onMouseLeave={e => { if (!host.state.alertOpen) e.currentTarget.style.background = 'rgba(255,255,255,0.15)'; }}
                aria-label="Alerts"
              >
                <Icon iconName="Ringer" style={{ fontSize: 20 }} />
                {host._unreadAlertCount() > 0 && (
                  <span style={{
                    position: 'absolute', top: 2, right: 2,
                    minWidth: 18, height: 18, borderRadius: 9,
                    background: '#ef4444', color: '#fff', fontSize: 10, fontWeight: 700,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    padding: '0 4px', border: `2px solid ${clay.accentDark}`,
                  }}>
                    {host._unreadAlertCount() > 99 ? '99+' : host._unreadAlertCount()}
                  </span>
                )}
              </button>

              {/* Alert Dropdown */}
              {host.state.alertOpen && (
                <div
                  data-alert-bell="true"
                  style={{
                    position: 'absolute', top: 'calc(100% + 10px)', right: 0,
                    width: phone ? Math.min(360, Math.max(290, viewportWidth - 24)) : 400, background: clay.surface, border: 'none',
                    borderRadius: clay.radiusCard, boxShadow: clay.shadowRaisedHover,
                    overflow: 'hidden', zIndex: 9999,
                  }}
                >
                  {/* Dropdown header */}
                  <div style={{
                    padding: '14px 18px', borderBottom: `1px solid ${clay.accentSoft}`,
                    background: clay.bg,
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <Icon iconName="Ringer" style={{ fontSize: 16, color: clay.accentDark }} />
                      <span style={{ fontSize: 15, fontWeight: 700, color: clay.text }}>Alerts</span>
                      {host._unreadAlertCount() > 0 && (
                        <span style={{ background: clay.pillDangerBg, color: clay.pillDangerText, borderRadius: 10, padding: '1px 8px', fontSize: 11, fontWeight: 700, boxShadow: clay.pillDangerShadow }}>
                          {host._unreadAlertCount()}
                        </span>
                      )}
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                      {([
                        ['dms', 'Vessels'],
                        ['crud', 'SPFx activity'],
                        ['email', 'Email alerts'],
                      ] as const).map(([category, label]) => (
                        <button
                          key={category}
                          onClick={() => host._setAlertCategory(category)}
                          style={{
                            padding: '4px 9px', borderRadius: 16, border: 'none',
                            background: host.state.alertCategory === category ? clay.accentGradient : clay.surfaceRaised,
                            color: host.state.alertCategory === category ? '#fff' : clay.textMuted,
                            fontSize: 11, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap',
                            boxShadow: host.state.alertCategory === category ? clay.shadowIcon : 'none',
                          }}
                        >{label}</button>
                      ))}
                      {(['all', 'unread'] as const).map(f => (
                        <button
                          key={f}
                          onClick={() => host._setAlertFilter(f)}
                          style={{
                            padding: '4px 12px', borderRadius: 16, border: 'none',
                            background: host.state.alertFilter === f ? clay.accentGradient : clay.surfaceRaised,
                            color: host.state.alertFilter === f ? '#fff' : clay.textMuted,
                            fontSize: 12, fontWeight: 600, cursor: 'pointer',
                            textTransform: 'capitalize',
                            boxShadow: host.state.alertFilter === f ? clay.shadowIcon : 'none',
                          }}
                        >{f}</button>
                      ))}
                      {host._unreadAlertCount() > 0 && (
                        <button
                          onClick={host._markAllAlertsRead}
                          style={{
                            padding: '4px 10px', borderRadius: 16, border: 'none',
                            background: clay.pillWarnBg, color: clay.pillWarnText,
                            fontSize: 11, fontWeight: 600, cursor: 'pointer', boxShadow: clay.pillWarnShadow,
                          }}
                        >Mark all read</button>
                      )}
                      <button
                        type="button"
                        onClick={() => host._openAlertsPage()}
                        title="Open full alerts page"
                        style={{ padding: '4px 9px', borderRadius: 16, border: 'none', background: clay.accentSoft, color: clay.accentDark, fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
                      >⛶ Maximize</button>
                    </div>
                  </div>

                  <div style={{ maxHeight: 420, overflowY: 'auto' }}>
                    {(() => {
                      const filtered = host.state.alertsList.filter(
                          a => (a.alert_category === 'crud' || a.alert_type === 'crud_operation' ? 'crud'
                          : a.alert_category === 'email' || a.alert_type === 'email_alert' ? 'email'
                          : 'dms'
                        ) === host.state.alertCategory && (host.state.alertFilter === 'all' || !a.read)
                      );
                      if (filtered.length === 0) {
                        return (
                          <div style={{ padding: '32px 20px', textAlign: 'center' }}>
                            <Icon iconName="CheckMark" style={{ fontSize: 32, color: '#10b981', display: 'block', margin: '0 auto 10px' }} />
                            <div style={{ fontSize: 14, fontWeight: 600, color: 'var(--vdms-text)', marginBottom: 4 }}>
                              {host.state.alertsList.length === 0 ? 'No alerts yet' : 'All caught up!'}
                            </div>
                            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>
                              {host.state.alertsList.length === 0
                                ? 'New SharePoint folder/vessel creations will appear here.'
                                : 'No unread alerts.'}
                            </div>
                          </div>
                        );
                      }
                      return (
                        <>
                          {filtered.map((alert: AlertItem) => {
                            const isAnomaly = alert.alert_type === 'vessel_unrecognised'
                              || alert.alert_type === 'file_outside_structure'
                              || alert.alert_type === 'subfolder_anomaly';

                            const iconName = alert.alert_type === 'vessel_provisioned' ? 'Ferry'
                              : alert.alert_type === 'vessel_deleted' ? 'Delete'
                              : alert.alert_type === 'document_deleted' ? 'Delete'
                              : alert.alert_type === 'vessel_unrecognised' ? 'Warning'
                              : alert.alert_type === 'file_outside_structure' ? 'PageSolid'
                              : alert.alert_type === 'subfolder_anomaly' ? 'FabricNewFolder'
                              : 'FabricNewFolder';

                            const iconBg = alert.alert_type === 'vessel_provisioned' ? '#dcfce7'
                              : alert.alert_type === 'vessel_deleted' ? '#fee2e2'
                              : alert.alert_type === 'document_deleted' ? '#fee2e2'
                              : alert.alert_type === 'vessel_unrecognised' ? '#fef3c7'
                              : alert.alert_type === 'file_outside_structure' ? '#f3e8ff'
                              : alert.alert_type === 'subfolder_anomaly' ? '#fff7ed'
                              : '#dbeafe';

                            const iconColor = alert.alert_type === 'vessel_provisioned' ? '#166534'
                              : alert.alert_type === 'vessel_deleted' ? '#991b1b'
                              : alert.alert_type === 'document_deleted' ? '#991b1b'
                              : alert.alert_type === 'vessel_unrecognised' ? '#92400e'
                              : alert.alert_type === 'file_outside_structure' ? '#6b21a8'
                              : alert.alert_type === 'subfolder_anomaly' ? '#9a3412'
                              : '#1e40af';

                            const typeLabel = alert.alert_type === 'vessel_provisioned' ? 'Provisioned'
                              : alert.alert_type === 'vessel_deleted' ? 'Deleted'
                              : alert.alert_type === 'document_deleted' ? 'Deleted'
                              : alert.alert_type === 'vessel_unrecognised' ? 'Unrecognised'
                              : alert.alert_type === 'file_outside_structure' ? 'File'
                              : alert.alert_type === 'subfolder_anomaly' ? 'Anomaly'
                              : 'New Folder';

                            return (
                              <div
                                key={alert.id}
                                onClick={() => host._openAlertsPage(alert.id)}
                                style={{
                                  padding: '12px 18px', borderBottom: '1px solid var(--vdms-border-soft)',
                                  background: alert.read ? 'transparent' : '#f8faff',
                                  cursor: 'pointer', transition: 'background 0.1s ease',
                                }}
                              >
                                <div style={{ display: 'flex', alignItems: 'flex-start', gap: 12 }}>
                                  <div style={{
                                    width: 36, height: 36, borderRadius: 10,
                                    background: iconBg, color: iconColor,
                                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                                    fontSize: 18, flexShrink: 0,
                                  }}>
                                    <Icon iconName={iconName} />
                                  </div>
                                  <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                                      <span style={{
                                        fontSize: 13, fontWeight: 600, color: 'var(--vdms-text)',
                                        whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                                        maxWidth: 180,
                                      }}>
                                        {alert.folder_name}
                                      </span>
                                      <span style={{
                                        fontSize: 9, fontWeight: 700, padding: '2px 7px', borderRadius: 10,
                                        background: iconBg, color: iconColor,
                                        textTransform: 'uppercase', letterSpacing: '0.4px', flexShrink: 0,
                                      }}>
                                        {typeLabel}
                                      </span>
                                      {!alert.read && (
                                        <span style={{ width: 7, height: 7, borderRadius: '50%', background: '#ef4444', flexShrink: 0 }} />
                                      )}
                                    </div>
                                    <div style={{ marginTop: 2, fontSize: 11, color: 'var(--vdms-text-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                      {alert.folder_path}
                                    </div>
                                    {(alert.vessel_name || alert.department) && (
                                      <div style={{ marginTop: 4, fontSize: 11, color: 'var(--vdms-text-faint)', display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                                        {alert.vessel_name && (
                                          <span style={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                                            <Icon iconName="Ferry" style={{ fontSize: 10 }} /> {alert.vessel_name}
                                          </span>
                                        )}
                                        {alert.department && alert.department !== 'All Departments' && (
                                          <span style={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                                            <Icon iconName="Org" style={{ fontSize: 10 }} /> {alert.department}
                                          </span>
                                        )}
                                      </div>
                                    )}
                                    <div style={{ marginTop: 4, fontSize: 10, color: 'var(--vdms-text-faint)' }}>
                                      {isAnomaly
                                        ? 'Detected in SharePoint Online — needs classification'
                                        : alert.alert_type === 'vessel_deleted'
                                          ? `Moved to Recycle Bin by ${alert.created_by_name || alert.created_by_email}`
                                          : alert.alert_type === 'document_deleted'
                                            ? `Moved to Recycle Bin by ${alert.created_by_name || alert.created_by_email}`
                                          : `Created by ${alert.created_by_name || alert.created_by_email}`}
                                      {alert.created_at && ` • ${new Date(alert.created_at).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}`}
                                    </div>
                                    {alert.alert_type === 'vessel_deleted' && (
                                      <button
                                        onClick={e => {
                                          e.stopPropagation();
                                          if (!alert.read) host._markAlertRead(alert.id);
                                          host._closeAlertBell();
                                          void host._goToView('recycle');
                                        }}
                                        style={{
                                          marginTop: 7, padding: '4px 12px', borderRadius: 6,
                                          border: '1px solid #fca5a5', background: '#fff5f5',
                                          color: '#dc2626', fontSize: 11, fontWeight: 600,
                                          cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4,
                                        }}
                                      >
                                        <Icon iconName="NavigateForward" style={{ fontSize: 10 }} />
                                        View in Recycle Bin
                                      </button>
                                    )}
                                    {alert.alert_type === 'document_deleted' && (
                                      <button onClick={e => { e.stopPropagation(); if (!alert.read) host._markAlertRead(alert.id); host._closeAlertBell(); void host._goToView('recycle'); }} style={{ marginTop: 7, padding: '4px 12px', borderRadius: 6, border: '1px solid #fca5a5', background: '#fff5f5', color: '#dc2626', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>
                                        <Icon iconName="NavigateForward" style={{ fontSize: 10 }} /> View in Recycle Bin
                                      </button>
                                    )}
                                    {alert.alert_type === 'vessel_unrecognised' && (() => {
                                      const anomaly = host.state.folderAnomalies.find(
                                        a => a.id === alert.anomaly_id ||
                                          (a.drive_item_id === alert.drive_item_id && a.name === alert.folder_name)
                                      ) || {
                                        id: alert.anomaly_id ?? Date.now(),
                                        drive_item_id: alert.drive_item_id || '',
                                        name: alert.folder_name,
                                        item_type: 'folder' as const,
                                        anomaly_type: 'vessel_level_unmatched' as const,
                                        department: alert.department,
                                        vessel_name: alert.vessel_name,
                                        spo_path: alert.spo_path || alert.folder_path,
                                        resolved: false,
                                        detected_at: alert.created_at,
                                      };
                                      return (
                                        <div style={{ marginTop: 7, display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                                          <button
                                            onClick={e => {
                                              e.stopPropagation();
                                              if (!alert.read) host._markAlertRead(alert.id);
                                              host._closeAlertBell();
                                              void host._goToView('vessels');
                                              host.setState({ spoClassifyDialog: { anomaly, provisioning: false, done: false, error: null } });
                                            }}
                                            style={{
                                              padding: '4px 10px', borderRadius: 6,
                                              border: '1px solid #bae6fd', background: '#e0f2fe',
                                              color: '#0369a1', fontSize: 11, fontWeight: 700,
                                              cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4,
                                            }}
                                          >
                                            🚢 Make it a Vessel
                                          </button>
                                          <button
                                            onClick={e => {
                                              e.stopPropagation();
                                              if (!alert.read) host._markAlertRead(alert.id);
                                              host._closeAlertBell();
                                              void host._goToView('vessels');
                                              host.setState({ spoClassifyDialog: { anomaly, provisioning: false, done: false, doneNormal: false, error: null } });
                                            }}
                                            style={{
                                              padding: '4px 10px', borderRadius: 6,
                                              border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface-alt)',
                                              color: 'var(--vdms-text-secondary)', fontSize: 11, fontWeight: 600,
                                              cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4,
                                            }}
                                          >
                                            📁 Normal Folder
                                          </button>
                                        </div>
                                      );
                                    })()}
                                    {(alert.alert_type === 'file_outside_structure' || alert.alert_type === 'subfolder_anomaly') && (
                                      <button
                                        onClick={e => {
                                          e.stopPropagation();
                                          if (!alert.read) host._markAlertRead(alert.id);
                                          host._closeAlertBell();
                                          void host._goToView('vessels');
                                          if (alert.alert_type === 'file_outside_structure') {
                                            // Build subfolder options from known rows for this vessel
                                            const normName = (s: string) => (s || '').trim().toLowerCase();
                                            const vName = alert.vessel_name;
                                            const subFolderOptions = vName
                                              ? Array.from(
                                                  new Map(
                                                    host.state.rows
                                                      .filter(r => normName(r.vesselName) === normName(vName) && r.canUpload)
                                                      .map(r => [r.groupKey, { label: r.subFolderPath, groupKey: r.groupKey, uploadFolderId: r.uploadFolderId, subFolderPath: r.subFolderPath }])
                                                  ).values()
                                                ).slice(0, 40)
                                              : [];
                                            host.setState({
                                              spoFileAlertDialog: {
                                                fileId: alert.drive_item_id || '',
                                                fileName: alert.folder_name,
                                                spoPath: alert.spo_path || alert.folder_path,
                                                vesselName: alert.vessel_name,
                                                subFolderOptions,
                                                moving: false,
                                                moved: false,
                                                error: null,
                                              },
                                            });
                                          }
                                        }}
                                        style={{
                                          marginTop: 7, padding: '4px 12px', borderRadius: 6,
                                          border: '1px solid #bfdbfe', background: '#eff6ff',
                                          color: clay.accentDark, fontSize: 11, fontWeight: 600,
                                          cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4,
                                        }}
                                      >
                                        <Icon iconName="NavigateForward" style={{ fontSize: 10 }} />
                                        {alert.alert_type === 'file_outside_structure' ? '📁 Review File' : 'Track in Vessels'}
                                      </button>
                                    )}
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </>
                      );
                    })()}
                  </div>
                </div>
              )}
            </div>
            {/* End Alert Bell */}

            {/* Avatar */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <button type="button" onClick={() => host._goToView('profile')} title="Open profile" aria-label="Open profile" style={{
                width: 44, height: 44, borderRadius: '50%',
                background: 'rgba(255,255,255,0.25)',
                border: '2.5px solid rgba(255,255,255,0.7)',
                color: '#fff', fontSize: 18, fontWeight: 900,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                flexShrink: 0, boxShadow: '0 4px 14px rgba(0,0,0,0.25)',
                cursor: 'pointer',
              }}>
                {userDisplayName.charAt(0).toUpperCase()}
              </button>
              <span style={{ fontSize: 16, fontWeight: 800, color: '#fff', letterSpacing: '0.2px', textShadow: '0 1px 6px rgba(0,0,0,0.2)' }}>{userDisplayName}</span>
            </div>
          </div>
        </div>

        {host.state.scanProgress.status !== 'idle' && (
          <div role="status" style={{ flexShrink: 0, padding: '7px 18px', background: host.state.scanProgress.status === 'failed' ? '#fef2f2' : '#eff6ff', borderBottom: '1px solid #bfdbfe', color: host.state.scanProgress.status === 'failed' ? '#b91c1c' : '#075985', display: 'flex', alignItems: 'center', gap: 10, fontSize: 12 }}>
            <span style={{ fontWeight: 700, whiteSpace: 'nowrap' }}>
              {host.state.scanProgress.status === 'running' ? 'Scanning files' : host.state.scanProgress.status === 'completed' ? 'Scan complete' : 'Scan failed'}
            </span>
            <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {host.state.scanProgress.title}
              {host.state.scanProgress.status === 'running' && host.state.scanProgress.recentFiles[0] ? ` · ${host.state.scanProgress.recentFiles[0]}` : ''}
            </span>
            {host.state.scanProgress.status === 'running' && <span style={{ fontWeight: 700, whiteSpace: 'nowrap' }}>{host.state.scanProgress.completed} / {host.state.scanProgress.total}</span>}
            {host.state.scanProgress.status === 'completed' && (
              <button type="button" onClick={() => host.setState({ scanProgress: { ...host.state.scanProgress, status: 'idle' } })} style={{ border: 0, background: 'transparent', color: '#075985', cursor: 'pointer', fontWeight: 700 }} aria-label="Dismiss scan status">Dismiss</button>
            )}
          </div>
        )}

        {/* ── Main Content ── */}
        <div style={{ flex: 1, overflowY: 'auto', padding: isWorkspaceFullScreen ? (phone ? 10 : 18) : (phone ? 12 : tabletOrBelow ? 18 : 32), background: isNight ? '#211812' : clay.bg, ['--vdms-content-pad' as any]: `${isWorkspaceFullScreen ? (phone ? 10 : 18) : (phone ? 12 : tabletOrBelow ? 18 : 32)}px` }}>
          {content}
        </div>

        {/* Right-side Document Preview Drawer */}
        {host._renderDocPreviewDrawer()}
      </div>
    </div>
    </ThemeProvider>
  );
}
