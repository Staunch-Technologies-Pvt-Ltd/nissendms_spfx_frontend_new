import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import type { AlertItem, FolderAnomalyItem, NormalFolderRecord } from '../types/ui';
import { MAIN_FOLDERS } from '../vesselFolderTemplate';

type AlertCategory = 'dms' | 'unclassified' | 'classified' | 'crud' | 'email';

const categoryLabels: Record<AlertCategory, string> = {
  dms: 'DMS folders',
  unclassified: 'SharePoint site Unclassified Items',
  classified: 'SharePoint Classified Items',
  crud: 'SPFx activity',
  email: 'Email alerts',
};

function categoryOf(alert: AlertItem): AlertCategory {
  if (alert.alert_type === 'vessel_unrecognised' || alert.alert_type === 'file_outside_structure') return 'unclassified';
  return alert.alert_category || (alert.alert_type === 'crud_operation' ? 'crud' : alert.alert_type === 'email_alert' ? 'email' : 'dms');
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

function alertForNormalFolder(item: NormalFolderRecord): AlertItem {
  return {
    id: `classified_${item.id ?? item.drive_item_id}`,
    drive_item_id: item.drive_item_id,
    folder_name: item.name,
    folder_path: item.spo_path,
    parent_folder_id: null,
    vessel_name: item.vessel_name,
    department: item.department,
    created_by_email: '',
    created_by_name: 'SharePoint Online',
    alert_type: 'folder_created',
    alert_category: 'classified',
    read: true,
    created_at: item.detected_at,
    anomaly_id: item.id ?? undefined,
    item_type: item.item_type,
    spo_path: item.spo_path,
  };
}

export function renderAlertsPage(host: VesselEmail): React.ReactElement {
  const category = host.state.alertCategory;
  const classifiedAlerts = host.state.normalFolders.map(alertForNormalFolder);
  const allAlerts = [...classifiedAlerts, ...host.state.alertsList];
  const alerts = allAlerts.filter(alert => categoryOf(alert) === category);
  const selected = host.state.selectedAlertId
    ? allAlerts.find(alert => alert.id === host.state.selectedAlertId)
    : alerts[0];
  const selectedClassified = category === 'classified' && selected
    ? host.state.normalFolders.find(item => `classified_${item.id ?? item.drive_item_id}` === selected.id)
    : undefined;

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
            {selectedClassified ? (
              <ClassifiedDetails host={host} item={selectedClassified} />
            ) : <dl style={{ margin: 0, display: 'grid', gap: 12, fontSize: 13 }}>
              <div><dt style={{ color: '#94a3b8' }}>Location / detail</dt><dd style={{ margin: '3px 0 0', color: '#334155', wordBreak: 'break-word' }}>{selected.folder_path || 'No detail provided'}</dd></div>
              <div><dt style={{ color: '#94a3b8' }}>Department</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{selected.department || 'Application'}</dd></div>
              <div><dt style={{ color: '#94a3b8' }}>Created by</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{selected.created_by_name || selected.created_by_email || 'System'}</dd></div>
              <div><dt style={{ color: '#94a3b8' }}>Time</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{selected.created_at ? new Date(selected.created_at).toLocaleString() : 'Recent'}</dd></div>
            </dl>}
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 20 }}>
              {!selected.read && <button type="button" onClick={() => host._markAlertRead(selected.id)} style={{ border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Mark as read</button>}
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

function ClassifiedDetails({ host, item }: { host: VesselEmail; item: NormalFolderRecord }): React.ReactElement {
  const [editing, setEditing] = React.useState(false);
  const [group, setGroup] = React.useState(item.department || MAIN_FOLDERS[0]?.name || '');
  const groupNode = MAIN_FOLDERS.find(folder => folder.name === group);
  const [category, setCategory] = React.useState(groupNode?.perVesselTree?.[0]?.name || '');
  const categoryNode = groupNode?.perVesselTree?.find(folder => folder.name === category);
  const [subCategory, setSubCategory] = React.useState(categoryNode?.children?.[0]?.name || '');
  const [busy, setBusy] = React.useState(false);
  const [message, setMessage] = React.useState('');

  React.useEffect(() => {
    const nextGroup = MAIN_FOLDERS.find(folder => folder.name === group);
    setCategory(nextGroup?.perVesselTree?.[0]?.name || '');
  }, [group]);
  React.useEffect(() => {
    const nextCategory = MAIN_FOLDERS.find(folder => folder.name === group)?.perVesselTree?.find(folder => folder.name === category);
    setSubCategory(nextCategory?.children?.[0]?.name || '');
  }, [group, category]);

  const anomaly: FolderAnomalyItem = {
    id: item.id ?? 0, drive_item_id: item.drive_item_id, name: item.name, item_type: item.item_type,
    anomaly_type: 'classified_normal', department: item.department, vessel_name: item.vessel_name,
    spo_path: item.spo_path, resolved: true, detected_at: item.detected_at,
  };
  const run = async (action: () => Promise<void>): Promise<void> => {
    setBusy(true); setMessage('');
    try { await action(); setMessage('Classification updated.'); setEditing(false); host._loadNormalFolders(); }
    catch (error: any) { setMessage(error?.message || 'Could not update classification.'); }
    finally { setBusy(false); }
  };

  return <div style={{ fontSize: 13 }}>
    <div style={{ padding: '12px 14px', background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 8, marginBottom: 14 }}>
      <strong style={{ color: '#166534' }}>Classification Status</strong><div style={{ marginTop: 4, color: '#15803d' }}>Classified</div>
    </div>
    <dl style={{ margin: 0, display: 'grid', gap: 12 }}>
      <div><dt style={{ color: '#94a3b8' }}>Current SharePoint Location</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{item.spo_path}</dd></div>
      <div><dt style={{ color: '#94a3b8' }}>DMS Location</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{item.spo_path}</dd></div>
      <div><dt style={{ color: '#94a3b8' }}>Classification Type</dt><dd style={{ margin: '3px 0 0', color: '#334155' }}>{item.vessel_name ? 'Vessel' : 'Folder'}</dd></div>
    </dl>
    <button type="button" onClick={() => setEditing(value => !value)} style={{ marginTop: 18, border: '1px solid #0f766e', background: '#ecfdf5', color: '#0f766e', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Update Classification</button>
    {editing && <div style={{ marginTop: 16, borderTop: '1px solid #e2e8f0', paddingTop: 16 }}>
      <p style={{ margin: '0 0 12px', color: '#475569' }}>This item is already classified. You can update how it is organized in the DMS.</p>
      <button type="button" disabled={busy} onClick={() => run(async () => { host.setState({ spoClassifyDialog: { anomaly, provisioning: false, done: false, error: null } }); })} style={{ width: '100%', marginBottom: 12, border: '1px solid #bae6fd', background: '#e0f2fe', color: '#0369a1', borderRadius: 7, padding: '9px 12px', cursor: 'pointer', fontWeight: 700 }}>🚢 Change to Vessel</button>
      <label style={{ display: 'block', color: '#475569', fontWeight: 700, marginBottom: 5 }}>Group</label>
      <select value={group} onChange={event => setGroup(event.target.value)} style={{ width: '100%', padding: 8, marginBottom: 10 }}><option value="">Select Group</option>{MAIN_FOLDERS.map(folder => <option key={folder.name} value={folder.name}>{folder.name}</option>)}</select>
      <label style={{ display: 'block', color: '#475569', fontWeight: 700, marginBottom: 5 }}>Category</label>
      <select value={category} onChange={event => setCategory(event.target.value)} style={{ width: '100%', padding: 8, marginBottom: 10 }}><option value="">Select Category</option>{(groupNode?.perVesselTree || []).map(folder => <option key={folder.name} value={folder.name}>{folder.name}</option>)}</select>
      <label style={{ display: 'block', color: '#475569', fontWeight: 700, marginBottom: 5 }}>Sub-category</label>
      <select value={subCategory} onChange={event => setSubCategory(event.target.value)} style={{ width: '100%', padding: 8, marginBottom: 10 }}><option value="">Select Sub-category</option>{(categoryNode?.children || []).map(folder => <option key={folder.name} value={folder.name}>{folder.name}</option>)}</select>
      <div style={{ padding: 10, background: '#f8fafc', borderRadius: 7, color: '#475569', marginBottom: 10 }}>{group || '[Group]'} / {category || '[Category]'} / {subCategory || '[Sub-category]'} / {item.name}</div>
      <button type="button" disabled={busy || !group || !category || !subCategory} onClick={() => run(async () => {
        const scope = item.vessel_name === 'Kaizen - Knowledge Bank' ? 'kaizen' : item.vessel_name === 'Common for all vessels' ? 'common' : 'vessels';
        await host._placeAnomalyInDmsCategory(anomaly, scope, item.vessel_name || '', group, [category, subCategory].filter(Boolean));
      })} style={{ width: '100%', border: 0, background: '#0f766e', color: '#fff', borderRadius: 7, padding: '9px 12px', cursor: 'pointer', fontWeight: 700 }}>Update DMS Location</button>
      <button type="button" disabled={busy} onClick={() => { setEditing(false); setMessage('Current location retained.'); }} style={{ width: '100%', marginTop: 8, border: '1px solid #cbd5e1', background: '#fff', color: '#475569', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontWeight: 700 }}>Keep Current Location</button>
      {message && <div style={{ marginTop: 10, color: message.startsWith('Could') ? '#b91c1c' : '#15803d' }}>{message}</div>}
    </div>}
  </div>;
}
