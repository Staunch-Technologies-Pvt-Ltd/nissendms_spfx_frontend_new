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
  ApprovalItem, NotificationItem, UserItem, AlertItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import { isMobileWidth, isTabletOrBelow } from '../responsive';

// Must match NAV_ACCENTS in Sidebar.tsx
const VIEW_ACCENTS: Record<string, string> = {
  dashboard:    '#FF2D55',
  list:         '#FF6B00',
  vessels:      '#F5C400',
  templates:    '#00C853',
  approvals:    '#00B0FF',
  reports:      '#3D5AFE',
  users:        '#AA00FF',
  settings:     '#F50057',
  bento_email:  '#FF3D00',
  email_notify: '#FF3D00',
  bento_compose:'#FF6D00',
  recycle:      '#00BCD4',
  archive:      '#6200EA',
};

// View → display label
const VIEW_LABELS: Record<string, string> = {
  dashboard: 'Dashboard',
  list: 'Documents',
  vessels: 'Vessels',
  templates: 'Templates',
  approvals: 'Approvals',
  reports: 'Reports',
  users: 'User Management',
  settings: 'Settings',
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
    /* Placeholder input text color fix */
    .vessel-dms-search::placeholder { color: rgba(255,255,255,0.55); }
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

    /* ── Animated Vessel Banner ── */
    @keyframes vesselShipMotion {
      0% {
        transform: translate(0px, 0px) rotate(0deg);
      }
      25% {
        transform: translate(2px, -3.5px) rotate(-1.2deg);
      }
      50% {
        transform: translate(3.5px, 0.5px) rotate(0.4deg);
      }
      75% {
        transform: translate(1.5px, 3px) rotate(1.2deg);
      }
      100% {
        transform: translate(0px, 0px) rotate(0deg);
      }
    }

    @keyframes vesselWaveScroll {
      0% {
        transform: translateX(0);
      }
      100% {
        transform: translateX(-400px);
      }
    }

    @keyframes vesselCloudDrift {
      0% {
        transform: translateX(0);
      }
      50% {
        transform: translateX(20px);
      }
      100% {
        transform: translateX(0);
      }
    }

    @keyframes vesselBowSprayPulse {
      0%, 100% {
        opacity: 0.4;
        transform: scale(0.95);
      }
      25% {
        opacity: 0.85;
        transform: scale(1.1);
      }
      75% {
        opacity: 0.35;
        transform: scale(0.9);
      }
    }

    .vessel-ship-animated {
      animation: vesselShipMotion 3.8s ease-in-out infinite;
      transform-origin: 200px 115px;
      will-change: transform;
    }

    .vessel-wave-back {
      animation: vesselWaveScroll 6.5s linear infinite;
      will-change: transform;
    }

    .vessel-wave-mid {
      animation: vesselWaveScroll 4s linear infinite;
      will-change: transform;
    }

    .vessel-wave-front {
      animation: vesselWaveScroll 3s linear infinite;
      will-change: transform;
    }

    .vessel-cloud-slow {
      animation: vesselCloudDrift 18s ease-in-out infinite;
      will-change: transform;
    }

    .vessel-cloud-fast {
      animation: vesselCloudDrift 12s ease-in-out infinite reverse;
      will-change: transform;
    }

    .vessel-bow-spray {
      animation: vesselBowSprayPulse 3.8s ease-in-out infinite;
      transform-origin: 390px 125px;
    }

    /* Accessibility: Respect prefers-reduced-motion */
    @media (prefers-reduced-motion: reduce) {
      .vessel-ship-animated,
      .vessel-wave-back,
      .vessel-wave-mid,
      .vessel-wave-front,
      .vessel-cloud-slow,
      .vessel-cloud-fast,
      .vessel-bow-spray {
        animation: none !important;
        transform: none !important;
      }
    }
  `;
}

export function renderLayout(host: VesselEmail, content: React.ReactElement): React.ReactElement {
  const userDisplayName = host.props.userDisplayName || 'Admin';
  const viewLabel = VIEW_LABELS[host.state.view] || host.state.view.replace(/_/g, ' ');
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const tabletOrBelow = isTabletOrBelow(viewportWidth);
  const phone = isMobileWidth(viewportWidth);

  injectFullScreenStyles();

  return (
    <div style={{
      display: 'flex', alignItems: 'stretch',
      height: '100vh', minHeight: '100vh',
      width: '100vw', overflow: 'hidden',
      background: '#F0F4F8',
      fontFamily: "'Segoe UI Variable', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 99999,
    }}>
      {host._renderSidebar()}

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0, overflow: 'hidden' }}>

        {/* ── Top Bar ── */}
        <div style={{
          height: 76, flexShrink: 0,
          background: VIEW_ACCENTS[host.state.view] || '#3D5AFE',
          boxShadow: `0 4px 24px ${VIEW_ACCENTS[host.state.view] || '#3D5AFE'}88`,
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
                }}
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
          </div>

          {/* Right controls */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {/* Search */}
            {!phone && <div style={{ position: 'relative', width: tabletOrBelow ? 180 : 260 }}>
              <input
                type="text"
                placeholder="Search..."
                className="vessel-dms-search"
                style={{
                  width: '100%', padding: '10px 18px 10px 44px', borderRadius: 28,
                  border: '2px solid rgba(255,255,255,0.4)',
                  fontSize: 15, background: 'rgba(255,255,255,0.18)',
                  outline: 'none', boxSizing: 'border-box',
                  color: '#fff', backdropFilter: 'blur(8px)',
                  fontWeight: 500,
                }}
                onFocus={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.28)'; e.currentTarget.style.borderColor = 'rgba(255,255,255,0.7)'; }}
                onBlur={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.18)'; e.currentTarget.style.borderColor = 'rgba(255,255,255,0.35)'; }}
              />
              <Icon iconName="Search" style={{ position: 'absolute', left: 16, top: '50%', transform: 'translateY(-50%)', color: 'rgba(255,255,255,0.85)', fontSize: 17 }} />
            </div>}

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
                    padding: '0 4px', border: '2px solid #0078D4',
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
                    width: phone ? Math.min(360, Math.max(290, viewportWidth - 24)) : 400, background: '#fff', border: '1px solid #e2e8f0',
                    borderRadius: 14, boxShadow: '0 12px 40px rgba(16,24,40,.22)',
                    overflow: 'hidden', zIndex: 9999,
                  }}
                >
                  {/* Dropdown header */}
                  <div style={{
                    padding: '14px 18px', borderBottom: '1px solid #f1f5f9',
                    background: 'linear-gradient(135deg, #f8faff, #eff6ff)',
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <Icon iconName="Ringer" style={{ fontSize: 16, color: '#0078d4' }} />
                      <span style={{ fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Alerts</span>
                      {host._unreadAlertCount() > 0 && (
                        <span style={{ background: '#ef4444', color: '#fff', borderRadius: 10, padding: '1px 8px', fontSize: 11, fontWeight: 700 }}>
                          {host._unreadAlertCount()}
                        </span>
                      )}
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                      {([
                        ['dms', 'DMS folders'],
                        ['unclassified', 'SharePoint site Unclassified Items'],
                        ['classified', 'SharePoint Classified Items'],
                        ['crud', 'SPFx activity'],
                        ['email', 'Email alerts'],
                      ] as const).map(([category, label]) => (
                        <button
                          key={category}
                          onClick={() => host._setAlertCategory(category)}
                          style={{
                            padding: '4px 9px', borderRadius: 16, border: 'none',
                            background: host.state.alertCategory === category ? '#0f766e' : '#f1f5f9',
                            color: host.state.alertCategory === category ? '#fff' : '#475569',
                            fontSize: 11, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap',
                          }}
                        >{label}</button>
                      ))}
                      {(['all', 'unread'] as const).map(f => (
                        <button
                          key={f}
                          onClick={() => host._setAlertFilter(f)}
                          style={{
                            padding: '4px 12px', borderRadius: 16, border: 'none',
                            background: host.state.alertFilter === f ? '#0078d4' : '#f1f5f9',
                            color: host.state.alertFilter === f ? '#fff' : '#475569',
                            fontSize: 12, fontWeight: 600, cursor: 'pointer',
                            textTransform: 'capitalize',
                          }}
                        >{f}</button>
                      ))}
                      {host._unreadAlertCount() > 0 && (
                        <button
                          onClick={host._markAllAlertsRead}
                          style={{
                            padding: '4px 10px', borderRadius: 16, border: 'none',
                            background: '#fef3c7', color: '#92400e',
                            fontSize: 11, fontWeight: 600, cursor: 'pointer',
                          }}
                        >Mark all read</button>
                      )}
                      <button
                        type="button"
                        onClick={() => host._openAlertsPage()}
                        title="Open full alerts page"
                        style={{ padding: '4px 9px', borderRadius: 16, border: 'none', background: '#e0f2fe', color: '#0369a1', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
                      >⛶ Maximize</button>
                    </div>
                  </div>

                  <div style={{ maxHeight: 420, overflowY: 'auto' }}>
                    {(() => {
                      const filtered = host.state.alertsList.filter(
                          a => (a.alert_type === 'vessel_unrecognised' ? 'unclassified' : a.alert_category || (
                          a.alert_type === 'crud_operation' ? 'crud' : a.alert_type === 'email_alert' ? 'email' : 'dms'
                        )) === host.state.alertCategory && (host.state.alertFilter === 'all' || !a.read)
                      );
                      if (filtered.length === 0) {
                        return (
                          <div style={{ padding: '32px 20px', textAlign: 'center' }}>
                            <Icon iconName="CheckMark" style={{ fontSize: 32, color: '#10b981', display: 'block', margin: '0 auto 10px' }} />
                            <div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', marginBottom: 4 }}>
                              {host.state.alertsList.length === 0 ? 'No alerts yet' : 'All caught up!'}
                            </div>
                            <div style={{ fontSize: 12, color: '#64748b' }}>
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
                                  padding: '12px 18px', borderBottom: '1px solid #f1f5f9',
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
                                        fontSize: 13, fontWeight: 600, color: '#0f172a',
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
                                    <div style={{ marginTop: 2, fontSize: 11, color: '#64748b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                      {alert.folder_path}
                                    </div>
                                    {(alert.vessel_name || alert.department) && (
                                      <div style={{ marginTop: 4, fontSize: 11, color: '#94a3b8', display: 'flex', gap: 10, flexWrap: 'wrap' }}>
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
                                    <div style={{ marginTop: 4, fontSize: 10, color: '#94a3b8' }}>
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
                                              border: '1px solid #e2e8f0', background: '#f8fafc',
                                              color: '#475569', fontSize: 11, fontWeight: 600,
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
                                          color: '#0078d4', fontSize: 11, fontWeight: 600,
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
              <div style={{
                width: 44, height: 44, borderRadius: '50%',
                background: 'rgba(255,255,255,0.25)',
                border: '2.5px solid rgba(255,255,255,0.7)',
                color: '#fff', fontSize: 18, fontWeight: 900,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                flexShrink: 0, boxShadow: '0 4px 14px rgba(0,0,0,0.25)',
              }}>
                {userDisplayName.charAt(0).toUpperCase()}
              </div>
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
        <div style={{ flex: 1, overflowY: 'auto', padding: phone ? 12 : tabletOrBelow ? 18 : 32, background: '#EEF2F7' }}>
          {content}
        </div>

        {/* Right-side Document Preview Drawer */}
        {host._renderDocPreviewDrawer()}
      </div>
    </div>
  );
}
