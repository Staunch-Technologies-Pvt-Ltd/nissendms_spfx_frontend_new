import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import type { AlertItem } from '../types/ui';

type AlertCategory = 'dms' | 'crud' | 'email';

const categoryLabels: Record<AlertCategory, string> = {
  dms: 'DMS folders',
  crud: 'SPFx activity',
  email: 'Email alerts',
};

function categoryOf(alert: AlertItem): AlertCategory {
  return alert.alert_category || (alert.alert_type === 'crud_operation' ? 'crud' : alert.alert_type === 'email_alert' ? 'email' : 'dms');
}

function typeLabel(alert: AlertItem): string {
  if (alert.alert_type === 'crud_operation') return 'CRUD operation';
  if (alert.alert_type === 'email_alert') return 'Email notification';
  if (alert.alert_type === 'vessel_deleted') return 'Deleted';
  if (alert.alert_type === 'vessel_provisioned') return 'Provisioned';
  if (alert.alert_type === 'file_outside_structure') return 'File outside structure';
  if (alert.alert_type === 'vessel_unrecognised') return 'Unknown folder';
  if (alert.alert_type === 'subfolder_anomaly') return 'Folder anomaly';
  return 'Folder update';
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
          <div style={{ color: '#64748b', fontSize: 12, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1 }}>Activity center</div>
          <h1 style={{ margin: '6px 0 6px', color: '#0f172a', fontSize: 30, lineHeight: 1.1 }}>Alerts</h1>
          <p style={{ margin: 0, color: '#64748b', fontSize: 14 }}>Detailed DMS, application, and email events from the top bar.</p>
        </div>
        <button type="button" onClick={() => host._goToView('list')} style={{ border: '1px solid #cbd5e1', background: '#fff', color: '#334155', borderRadius: 8, padding: '8px 14px', cursor: 'pointer', fontWeight: 600 }}>
          <Icon iconName="ChevronLeft" /> Back
        </button>
      </div>

      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 20 }}>
        {(Object.keys(categoryLabels) as AlertCategory[]).map(key => (
          <button key={key} type="button" onClick={() => { host._setAlertCategory(key); host.setState({ selectedAlertId: null }); }} style={{ border: '1px solid #cbd5e1', background: category === key ? '#0f766e' : '#fff', color: category === key ? '#fff' : '#334155', borderRadius: 8, padding: '10px 16px', cursor: 'pointer', fontWeight: 700 }}>
            {categoryLabels[key]}
            <span style={{ marginLeft: 8, opacity: 0.75 }}>{host.state.alertsList.filter(alert => categoryOf(alert) === key).length}</span>
          </button>
        ))}
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: selected ? 'minmax(0, 1.4fr) minmax(300px, .8fr)' : '1fr', gap: 18, alignItems: 'start' }}>
        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, overflow: 'hidden' }}>
          <div style={{ padding: '14px 18px', borderBottom: '1px solid #e2e8f0', fontWeight: 700, color: '#0f172a' }}>{categoryLabels[category]}</div>
          {alerts.length === 0 ? (
            <div style={{ padding: 48, textAlign: 'center', color: '#64748b' }}>No alerts in this category.</div>
          ) : alerts.map(alert => (
            <button key={alert.id} type="button" onClick={() => host._openAlertsPage(alert.id)} style={{ width: '100%', textAlign: 'left', border: 0, borderBottom: '1px solid #f1f5f9', background: selected?.id === alert.id ? '#eff6ff' : '#fff', padding: '15px 18px', cursor: 'pointer' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
                <strong style={{ color: '#0f172a', fontSize: 14 }}>{alert.folder_name}</strong>
                <span style={{ color: alert.read ? '#94a3b8' : '#dc2626', fontSize: 11, fontWeight: 700 }}>{alert.read ? 'Read' : 'Unread'}</span>
              </div>
              <div style={{ color: '#64748b', fontSize: 12, marginTop: 4 }}>{typeLabel(alert)} · {alert.created_at ? new Date(alert.created_at).toLocaleString() : 'Recent'}</div>
              <div style={{ color: '#94a3b8', fontSize: 11, marginTop: 5, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{alert.folder_path}</div>
            </button>
          ))}
        </section>

        {selected && (
          <aside style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, padding: 22, position: 'sticky', top: 20 }}>
            <div style={{ color: '#0f766e', fontSize: 11, fontWeight: 800, textTransform: 'uppercase', letterSpacing: .8 }}>{typeLabel(selected)}</div>
            <h2 style={{ margin: '8px 0 14px', color: '#0f172a', fontSize: 20 }}>{selected.folder_name}</h2>
            <dl style={{ margin: 0, display: 'grid', gap: 12, fontSize: 13 }}>
              <div><dt style={{ color: '#94a3b8' }}>Location / detail</dt><dd style={{ margin: '3px 0 0', color: '#334155', wordBreak: 'break-word' }}>{selected.folder_path || 'No detail provided'}</dd></div>
              <div><dt style={{ color: '#94a3b8' }}>Department</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{selected.department || 'Application'}</dd></div>
              <div><dt style={{ color: '#94a3b8' }}>Created by</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{selected.created_by_name || selected.created_by_email || 'System'}</dd></div>
              <div><dt style={{ color: '#94a3b8' }}>Time</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{selected.created_at ? new Date(selected.created_at).toLocaleString() : 'Recent'}</dd></div>
            </dl>
            {!selected.read && <button type="button" onClick={() => host._markAlertRead(selected.id)} style={{ marginTop: 20, border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Mark as read</button>}
          </aside>
        )}
      </div>
    </div>
  );
}
