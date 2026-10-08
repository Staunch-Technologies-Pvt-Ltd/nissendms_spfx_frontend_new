import * as React from 'react';
import * as ReactDOM from 'react-dom';
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
import { injectFuturisticTheme } from '../futuristicTheme';
import { ShipTransitions } from './ShipScenes';
import { injectRefreshTheme } from '../refreshTheme';

// View → display label
const VIEW_LABELS: Record<string, string> = {
  dashboard: 'Dashboard',
  list: 'Vessel DMS',
  vessels: 'Vessels',
  templates: 'Templates',
  migration: 'Migration Assistant',
  users: 'User Management',
  settings: 'Settings',
  profile: 'Profile',
  alerts: 'Notifications',
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
       the app root below. Night values use the Maritime deep-sea palette (see deepHarborNightTheme / the night
       input styles further down) so nothing clashes. */
    [data-vessel-theme="light"] {
      --vdms-surface: #ffffff;
      --vdms-surface-alt: #f6f8fb;
      --vdms-border: #e2e8f0;
      --vdms-border-soft: #eef1f5;
      --vdms-text: #101b2d;
      --vdms-text-secondary: #475569;
      --vdms-text-muted: #64748b;
      --vdms-text-faint: #94a3b8;
      --vdms-toggle-active-bg: #0b2a4a;
      --vdms-toggle-active-text: #ffffff;
      /* clay.* tokens (clayTheme.ts) */
      --clay-bg: #eef2f7;
      --clay-surface: #ffffff;
      --clay-surface-raised: #f6f8fb;
      --clay-surface-hover: #eaf0f6;
      --clay-text: #101b2d;
      --clay-text-muted: #5b6b7f;
      --clay-accent-soft: #dceef2;
      --clay-accent-soft-hover: #c7e4ea;
      --clay-icon-bg: linear-gradient(150deg, #e3f2f4, #cfe8ec);
      --clay-shadow-raised: 0 1px 2px rgba(16,27,45,0.05), 0 4px 12px rgba(16,27,45,0.06);
      --clay-shadow-raised-hover: 0 2px 4px rgba(16,27,45,0.06), 0 8px 20px rgba(16,27,45,0.10);
      --clay-shadow-button: 0 1px 2px rgba(16,27,45,0.08), 0 4px 10px rgba(14,116,144,0.25);
      --clay-shadow-icon: inset 0 0 0 1px rgba(14,116,144,0.14);
      --clay-pill-active-bg: #dcfce7;
      --clay-pill-active-text: #15803d;
      --clay-pill-active-shadow: none;
      --clay-pill-warn-bg: #fef3c7;
      --clay-pill-warn-text: #b45309;
      --clay-pill-warn-shadow: none;
      --clay-pill-danger-bg: #fee2e2;
      --clay-pill-danger-text: #b91c1c;
      --clay-pill-danger-shadow: none;
    }
    [data-vessel-theme="night"] {
      --vdms-surface: #101d2e;
      --vdms-surface-alt: #152536;
      --vdms-border: rgba(148,178,204,0.18);
      --vdms-border-soft: rgba(148,178,204,0.1);
      --vdms-text: #e7eef5;
      --vdms-text-secondary: #aebfd1;
      --vdms-text-muted: #8ca0b5;
      --vdms-text-faint: #71869b;
      --vdms-toggle-active-bg: #2dd4bf;
      --vdms-toggle-active-text: #04211d;
      /* clay.* tokens (clayTheme.ts) — night values */
      --clay-bg: #0a1626;
      --clay-surface: #101d2e;
      --clay-surface-raised: #182536;
      --clay-surface-hover: #202f42;
      --clay-text: #e7eef5;
      --clay-text-muted: #8ca0b5;
      --clay-accent-soft: #113b38;
      --clay-accent-soft-hover: #15473f;
      --clay-icon-bg: linear-gradient(150deg, #113b38, #15473f);
      --clay-shadow-raised: 0 1px 2px rgba(0,0,0,0.3), 0 4px 14px rgba(0,0,0,0.35);
      --clay-shadow-raised-hover: 0 2px 4px rgba(0,0,0,0.35), 0 8px 22px rgba(0,0,0,0.4);
      --clay-shadow-button: 0 1px 2px rgba(0,0,0,0.35), 0 4px 12px rgba(45,212,191,0.2);
      --clay-shadow-icon: inset 0 0 0 1px rgba(45,212,191,0.18);
      --clay-pill-active-bg: rgba(34,197,94,0.16);
      --clay-pill-active-text: #86efac;
      --clay-pill-active-shadow: none;
      --clay-pill-warn-bg: rgba(245,158,11,0.16);
      --clay-pill-warn-text: #fcd34d;
      --clay-pill-warn-shadow: none;
      --clay-pill-danger-bg: rgba(239,68,68,0.16);
      --clay-pill-danger-text: #fca5a5;
      --clay-pill-danger-shadow: none;
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
      background: rgba(3,20,35,0.6) !important;
      color: #eaf6fd !important;
      border-color: rgba(140,210,240,0.28) !important;
      color-scheme: dark;
    }
    [data-vessel-theme="night"] input::placeholder,
    [data-vessel-theme="night"] textarea::placeholder {
      color: #7f9bb0 !important;
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
      outline: 2px solid rgba(10,126,168,0.7) !important;
      outline-offset: 2px;
      box-shadow: 0 0 0 4px rgba(10,126,168,0.16) !important;
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
      outline: 3px solid rgba(10,126,168,0.55);
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
            <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 14 }} />
          </button>
        </div>
      ))}
    </div>
  );
}

export function renderLayout(host: VesselEmail, content: React.ReactElement): React.ReactElement {
  const userDisplayName = host.props.userDisplayName || 'Admin';
  const isNight = host.state.themeMode === 'night';
  const classicOpen = host.state.classicSiteOpen;
  const isWorkspaceFullScreen = host.state.fullScreenWorkspace && host.state.view === 'list' && !classicOpen;
  const viewLabel = VIEW_LABELS[host.state.view] || host.state.view.replace(/_/g, ' ');
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const tabletOrBelow = isTabletOrBelow(viewportWidth);
  const phone = isMobileWidth(viewportWidth);

  injectFullScreenStyles();
  injectFuturisticTheme();
  injectRefreshTheme();

  // Settings → Color Management (color_settings_api.py) recolors the app at
  // runtime; the clay.* tokens above already pick that up via CSS custom
  // properties (applyColorTheme), and Fluent's own theme is rebuilt here
  // from the same saved colors so its handful of stock controls match.
  const colorSet = isNight ? host.state.colorTheme.night : host.state.colorTheme.light;
  const fluentTheme = buildDeepHarborTheme(colorSet, isNight);

  // Rendered via a portal straight onto <body>: the SharePoint modern-page
  // shell wraps the web part zone in ancestor elements that establish their
  // own containing block for `position: fixed` (a known SPFx/SharePoint
  // quirk), so this overlay's `inset: 0` was being measured against that
  // ancestor's box instead of the real viewport — leaving a stray strip of
  // page chrome visible above the app on every page. Mounting on <body>
  // (never transformed by SharePoint) makes `fixed` pin to the true
  // viewport. React event handling is unaffected — portals still bubble
  // through the React tree, not the DOM tree.
  return ReactDOM.createPortal((
    <ThemeProvider theme={fluentTheme}>
    <div className="vessel-dms-app" style={{
      display: 'flex', alignItems: 'stretch',
      height: '100vh', minHeight: '100vh',
      width: '100vw', overflow: 'hidden',
      background: clay.bg,
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 99999,
    }} data-vessel-theme={host.state.themeMode}>
      <DeletionToastLayer host={host} />
      <ShipTransitions host={host} />
      {!isWorkspaceFullScreen && host._renderSidebar()}

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0, overflow: 'hidden' }}>

        {/* ── Top Bar ── */}
        <div className="vessel-dms-topbar" style={{
          height: 76, flexShrink: 0,
          background: 'var(--vdms-surface)',
          borderBottom: '1px solid var(--vdms-border)',
          boxShadow: '0 1px 2px rgba(16,27,45,0.04)',
          transition: 'background 0.3s ease, box-shadow 0.3s ease',
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
                  border: '1px solid var(--vdms-line-strong)',
                  background: 'var(--vdms-field)',
                  color: 'var(--vdms-text)',
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  cursor: 'pointer',
                  flexShrink: 0,
                  transition: 'all 0.15s ease',
                }}
                onMouseEnter={e => { e.currentTarget.style.background = 'var(--vdms-glass-strong)'; }}
                onMouseLeave={e => { e.currentTarget.style.background = 'var(--vdms-field)'; }}
                aria-label="Open navigation"
              >
                <Icon iconName="GlobalNavButton" style={{ fontSize: 18 }} />
              </button>
            )}
            {!phone && <span className="vdms-crumb" style={{ color: 'var(--vdms-glass-strong)', fontSize: 16, fontWeight: 600, letterSpacing: '0.2px' }}>NKS DOCMAN</span>}
            <Icon iconName="ChevronRight" className="vdms-crumb" style={{ fontSize: 14, color: 'var(--vdms-glass-strong)' }} />
            <span style={{ color: 'var(--vdms-text)', fontSize: phone ? 18 : 26, fontWeight: 700, letterSpacing: '-0.02em', fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}>
              {viewLabel}
            </span>
          </div>

          {/* Right controls */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {/* Traditional SharePoint site is opened from the sidebar ("Documents", under Vessel DMS);
                while it is showing, this Back button returns to the app. */}
            {classicOpen && (
              <button
                type="button"
                onClick={host._closeClassicSite}
                title="Back to NKS DOCMAN"
                aria-label="Back to NKS DOCMAN"
                style={{
                  height: 42, padding: phone ? '0 12px' : '0 16px', borderRadius: 12,
                  border: '1px solid var(--vdms-line-strong)',
                  background: 'var(--vdms-field)',
                  color: 'var(--vdms-text)', display: 'inline-flex', alignItems: 'center', gap: 8,
                  fontSize: 14, fontWeight: 700, cursor: 'pointer', boxShadow: '0 3px 10px rgba(0,0,0,0.15)',
                }}
              >
                <Icon iconName="Back" style={{ fontSize: 16 }} />
                {!phone && 'Back'}
              </button>
            )}
            {host.state.view === 'list' && !classicOpen && (
              <button
                type="button"
                onClick={host._toggleFullScreenWorkspace}
                title={isWorkspaceFullScreen ? 'Exit full-screen workspace' : 'Open full-screen workspace'}
                aria-label={isWorkspaceFullScreen ? 'Exit full-screen workspace' : 'Open full-screen workspace'}
                style={{
                  width: 42, height: 42, borderRadius: 12,
                  border: '1px solid var(--vdms-line-strong)',
                  background: isWorkspaceFullScreen ? 'var(--vdms-glass-strong)' : 'var(--vdms-field)',
                  color: 'var(--vdms-text)', display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
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
                  background: host.state.alertOpen ? 'var(--vdms-glass-strong)' : 'var(--vdms-field)',
                  color: 'var(--vdms-text)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  cursor: 'pointer', position: 'relative',
                  transition: 'background 0.15s ease',
                  boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
                }}
                onMouseEnter={e => { if (!host.state.alertOpen) e.currentTarget.style.background = 'var(--vdms-glass-strong)'; }}
                onMouseLeave={e => { if (!host.state.alertOpen) e.currentTarget.style.background = 'var(--vdms-field)'; }}
                aria-label="Alerts"
              >
                <Icon iconName="Ringer" style={{ fontSize: 20 }} />
                {host._unreadAlertCount() > 0 && (
                  <span style={{
                    position: 'absolute', top: 2, right: 2,
                    minWidth: 18, height: 18, borderRadius: 9,
                    background: '#ef4444', color: '#fff', fontSize: 10, fontWeight: 700,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    padding: '0 4px', border: '2px solid var(--vdms-glass-strong)',
                  }}>
                    {host._unreadAlertCount() > 99 ? '99+' : host._unreadAlertCount()}
                  </span>
                )}
              </button>

            </div>
            {/* End Alert Bell */}

            {/* Avatar */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <button type="button" onClick={() => host._goToView('profile')} title="Open profile" aria-label="Open profile" style={{
                width: 44, height: 44, borderRadius: '50%',
                background: 'var(--vdms-glass-strong)',
                border: '2.5px solid var(--vdms-line-strong)',
                color: 'var(--vdms-text)', fontSize: 18, fontWeight: 900,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                flexShrink: 0, boxShadow: '0 4px 14px rgba(0,0,0,0.25)',
                cursor: 'pointer',
              }}>
                {userDisplayName.charAt(0).toUpperCase()}
              </button>
              <span style={{ fontSize: 16, fontWeight: 800, color: 'var(--vdms-text)', letterSpacing: '0.2px' }}>{userDisplayName}</span>
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
        <div style={{ flex: 1, overflowY: 'auto', padding: isWorkspaceFullScreen ? (phone ? 10 : 18) : (phone ? 12 : tabletOrBelow ? 18 : 24), background: 'transparent', ['--vdms-content-pad' as any]: `${isWorkspaceFullScreen ? (phone ? 10 : 18) : (phone ? 12 : tabletOrBelow ? 18 : 24)}px` }}>
          {/* The app stays mounted (hidden) while the classic site is shown, so Back restores it as-is. */}
          <div style={{ display: classicOpen ? 'none' : 'block' }}>{content}</div>
          {classicOpen && (
            <div style={{ display: 'flex', flexDirection: 'column', height: '100%', minHeight: 480, border: '1px solid var(--vdms-border)', borderRadius: 12, overflow: 'hidden', background: 'var(--vdms-surface)' }}>
              <div style={{ flexShrink: 0, display: 'flex', alignItems: 'center', gap: 10, padding: '6px 12px', borderBottom: '1px solid var(--vdms-border)', fontSize: 12, color: 'var(--vdms-text-muted)' }}>
                <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>Traditional SharePoint site — {host.state.classicSiteUrl}</span>
                <a href={host.state.classicSiteUrl} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--clay-accent, #0e7490)', fontWeight: 700, whiteSpace: 'nowrap', textDecoration: 'none' }}>
                  Open in new tab
                </a>
              </div>
              <iframe
                title="SharePoint Documents library"
                src={host.state.classicSiteUrl}
                style={{ flex: 1, width: '100%', border: 0, background: '#fff' }}
              />
            </div>
          )}
        </div>

        {/* Right-side Document Preview Drawer */}
        {host._renderDocPreviewDrawer()}
      </div>
    </div>
    </ThemeProvider>
  ), document.body);
}
