import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import type { AlertItem, FolderAnomalyItem } from '../types/ui';
import { clay } from '../clayTheme';

type AlertCategory = 'dms' | 'crud' | 'email';

const categoryLabels: Record<AlertCategory, string> = {
  dms: 'Vessels',
  crud: 'SPFx activity',
  email: 'Email alerts',
};

function categoryOf(alert: AlertItem): AlertCategory {
  return alert.alert_category === 'crud' || alert.alert_type === 'crud_operation'
    ? 'crud'
    : alert.alert_category === 'email' || alert.alert_type === 'email_alert'
      ? 'email'
      : 'dms';
}

function typeLabel(alert: AlertItem): string {
  if (alert.alert_type === 'crud_operation') return 'CRUD operation';
  if (alert.alert_type === 'email_alert') return 'Email notification';
  if (alert.alert_type === 'vessel_deleted') return 'Deleted';
  if (alert.alert_type === 'document_deleted') return 'Document deleted';
  if (alert.alert_type === 'vessel_provisioned') return 'Provisioned';
  if (alert.alert_type === 'file_outside_structure') return 'File outside structure';
  if (alert.alert_type === 'vessel_unrecognised') return 'Unknown folder';
  if (alert.alert_type === 'subfolder_anomaly') return 'Folder anomaly';
  return 'Folder update';
}

function anomalyForAlert(alert: AlertItem): FolderAnomalyItem {
  return {
    id: alert.anomaly_id ?? Date.now(),
    drive_item_id: alert.drive_item_id || '',
    name: alert.folder_name,
    item_type: alert.item_type || 'folder',
    anomaly_type: 'vessel_level_unmatched',
    department: alert.department,
    vessel_name: alert.vessel_name,
    spo_path: alert.spo_path || alert.folder_path,
    resolved: false,
    detected_at: alert.created_at,
  };
}

export function renderAlertsPage(host: VesselEmail): React.ReactElement {
  const category = host.state.alertCategory;
  const alerts = host.state.alertsList.filter(alert => categoryOf(alert) === category);
  const selected = host.state.selectedAlertId
    ? host.state.alertsList.find(alert => alert.id === host.state.selectedAlertId)
    : alerts[0];

  return (
    <div style={{ maxWidth: 1180, margin: '0 auto', padding: '32px 36px', width: '100%', boxSizing: 'border-box' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 20, marginBottom: 24 }}>
        <div>
          <div style={{ color: clay.textMuted, fontSize: 12, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1 }}>Activity center</div>
          <h1 style={{ margin: '6px 0 6px', color: clay.text, fontSize: 30, lineHeight: 1.1 }}>Alerts</h1>
          <p style={{ margin: 0, color: clay.textMuted, fontSize: 14 }}>Detailed DMS, application, and email events from the top bar.</p>
        </div>
        <button type="button" onClick={() => host._goToView('list')} style={{ border: 'none', background: clay.surface, color: clay.text, borderRadius: clay.radiusButton, padding: '8px 14px', cursor: 'pointer', fontWeight: 600, boxShadow: clay.shadowRaised }}>
          <Icon iconName="ChevronLeft" /> Back
        </button>
      </div>

      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 20 }}>
        {(Object.keys(categoryLabels) as AlertCategory[]).map(key => (
          <button key={key} type="button" onClick={() => { host._setAlertCategory(key); host.setState({ selectedAlertId: null }); }} style={{ border: 'none', background: category === key ? clay.accentGradient : clay.surface, color: category === key ? '#fff' : clay.text, borderRadius: clay.radiusButton, padding: '10px 16px', cursor: 'pointer', fontWeight: 700, boxShadow: category === key ? clay.shadowIcon : clay.shadowRaised }}>
            {categoryLabels[key]}
            <span style={{ marginLeft: 8, opacity: 0.75 }}>{host.state.alertsList.filter(alert => categoryOf(alert) === key).length}</span>
          </button>
        ))}
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: selected ? 'minmax(0, 1.4fr) minmax(300px, .8fr)' : '1fr', gap: 18, alignItems: 'start' }}>
        <section style={{ background: clay.surface, border: 'none', borderRadius: clay.radiusCard, overflow: 'hidden', boxShadow: clay.shadowRaised }}>
          <div style={{ padding: '14px 18px', borderBottom: `1px solid ${clay.accentSoft}`, fontWeight: 700, color: clay.text }}>{categoryLabels[category]}</div>
          {alerts.length === 0 ? (
            <div style={{ padding: 48, textAlign: 'center', color: clay.textMuted }}>No alerts in this category.</div>
          ) : alerts.map(alert => (
            <button key={alert.id} type="button" onClick={() => host._openAlertsPage(alert.id)} style={{ width: '100%', textAlign: 'left', border: 0, borderBottom: `1px solid ${clay.accentSoft}`, background: selected?.id === alert.id ? clay.bg : 'transparent', padding: '15px 18px', cursor: 'pointer' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
                <strong style={{ color: clay.text, fontSize: 14 }}>{alert.folder_name}</strong>
                <span style={{ color: alert.read ? clay.textMuted : '#dc2626', fontSize: 11, fontWeight: 700 }}>{alert.read ? 'Read' : 'Unread'}</span>
              </div>
              <div style={{ color: clay.textMuted, fontSize: 12, marginTop: 4 }}>{typeLabel(alert)} · {alert.created_at ? new Date(alert.created_at).toLocaleString() : 'Recent'}</div>
              <div style={{ color: clay.textMuted, fontSize: 11, marginTop: 5, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{alert.folder_path}</div>
              {(alert.alert_type === 'vessel_deleted' || alert.alert_type === 'document_deleted') && (
                <div style={{ color: clay.textMuted, fontSize: 11, marginTop: 3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  Deleted by {alert.created_by_name || alert.created_by_email || 'Unknown user'}{alert.site_name ? ` · ${alert.site_name}` : ''}
                </div>
              )}
            </button>
          ))}
        </section>

        {selected && (
          <aside style={{ background: clay.surface, border: 'none', borderRadius: clay.radiusCard, padding: 22, position: 'sticky', top: 20, boxShadow: clay.shadowRaised }}>
            <div style={{ color: clay.accentDark, fontSize: 11, fontWeight: 800, textTransform: 'uppercase', letterSpacing: .8 }}>{typeLabel(selected)}</div>
            <h2 style={{ margin: '8px 0 14px', color: clay.text, fontSize: 20 }}>{selected.folder_name}</h2>
            {(() => {
              const isDeletion = selected.alert_type === 'vessel_deleted' || selected.alert_type === 'document_deleted';
              return <dl style={{ margin: 0, display: 'grid', gap: 12, fontSize: 13 }}>
                <div><dt style={{ color: clay.textMuted }}>{isDeletion ? 'Original path' : 'Location / detail'}</dt><dd style={{ margin: '3px 0 0', color: clay.text, wordBreak: 'break-word' }}>{selected.folder_path || 'No detail provided'}</dd></div>
                {isDeletion && <div><dt style={{ color: clay.textMuted }}>Site</dt><dd style={{ margin: '3px 0 0', color: clay.text }}>{selected.site_name || 'Unknown site'}</dd></div>}
                <div><dt style={{ color: clay.textMuted }}>Department</dt><dd style={{ margin: '3px 0 0', color: clay.text }}>{selected.department || 'Application'}</dd></div>
                <div><dt style={{ color: clay.textMuted }}>{isDeletion ? 'Deleted by' : 'Created by'}</dt><dd style={{ margin: '3px 0 0', color: clay.text }}>{selected.created_by_name || selected.created_by_email || (isDeletion ? 'Unknown user' : 'System')}</dd></div>
                {isDeletion && selected.reason && <div><dt style={{ color: clay.textMuted }}>Reason</dt><dd style={{ margin: '3px 0 0', color: clay.text }}>{selected.reason}</dd></div>}
                {isDeletion && selected.item_type && <div><dt style={{ color: clay.textMuted }}>Item type</dt><dd style={{ margin: '3px 0 0', color: clay.text, textTransform: 'capitalize' }}>{selected.vessel_name && selected.alert_type === 'vessel_deleted' ? 'Vessel' : selected.item_type}</dd></div>}
                <div><dt style={{ color: clay.textMuted }}>Time</dt><dd style={{ margin: '3px 0 0', color: clay.text }}>{selected.created_at ? new Date(selected.created_at).toLocaleString() : 'Recent'}</dd></div>
              </dl>;
            })()}
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 20 }}>
              {!selected.read && <button type="button" onClick={() => host._markAlertRead(selected.id)} style={{ border: 'none', background: clay.accentSoft, color: clay.accentDark, borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Mark as read</button>}
              {(selected.alert_type === 'vessel_unrecognised' || selected.alert_type === 'file_outside_structure') && (
                <>
                  <button type="button" onClick={() => {
                    host._markAlertRead(selected.id);
                    host.setState({ spoClassifyDialog: { anomaly: anomalyForAlert(selected), provisioning: false, done: false, error: null } });
                  }} style={{ border: '1px solid #bae6fd', background: '#e0f2fe', color: '#0369a1', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Classify</button>
                  <button type="button" onClick={() => host._moveAnomalyToRecycleBin(anomalyForAlert(selected))} style={{ border: '1px solid #fca5a5', background: '#fff5f5', color: '#dc2626', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Move to Recycle Bin</button>
                </>
              )}
            </div>
          </aside>
        )}
      </div>
    </div>
  );
}

