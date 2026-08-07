import * as React from 'react';
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
  ApprovalItem, NotificationItem, UserItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';

export function renderNotificationsPage(host: VesselEmail): React.ReactElement {
    const { notificationsList, notificationFilter } = host.state;
    const filtered = notificationsList.filter(n => notificationFilter === 'all' || !n.read);

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Notifications</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Stay updated with important alerts and reminders.</p>
          </div>
          <button
            onClick={() => {
              const updated = notificationsList.map(n => ({ ...n, read: true }));
              host.setState({ notificationsList: updated });
            }}
            style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
          >
            Mark all as read
          </button>
        </div>

        {/* Filter buttons */}
        <div style={{ display: 'flex', gap: 8 }}>
          <button
            onClick={() => host.setState({ notificationFilter: 'all' })}
            style={{
              border: 'none', background: notificationFilter === 'all' ? '#0078d4' : '#f1f5f9',
              color: notificationFilter === 'all' ? '#fff' : '#475569', borderRadius: 16,
              padding: '4px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            All ({notificationsList.length})
          </button>
          <button
            onClick={() => host.setState({ notificationFilter: 'unread' })}
            style={{
              border: 'none', background: notificationFilter === 'unread' ? '#0078d4' : '#f1f5f9',
              color: notificationFilter === 'unread' ? '#fff' : '#475569', borderRadius: 16,
              padding: '4px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            Unread ({notificationsList.filter(n => !n.read).length})
          </button>
        </div>

        {/* Notification list */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {filtered.map(notif => (
            <div
              key={notif.id}
              style={{
                background: notif.read ? '#fff' : '#f0f9ff',
                borderRadius: 8, border: '1px solid #e2e8f0', padding: '14px 18px',
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <div style={{
                  width: 32, height: 32, borderRadius: '50%',
                  background: notif.type === 'alert' ? '#fde7e9' : notif.type === 'warning' ? '#fff4ce' : notif.type === 'success' ? '#dff6dd' : '#e1efff',
                  color: notif.type === 'alert' ? '#a4262c' : notif.type === 'warning' ? '#8a5700' : notif.type === 'success' ? '#107c10' : '#0078d4',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700,
                }}>
                  {notif.type === 'alert' ? '⚠' : notif.type === 'warning' ? '🔔' : notif.type === 'success' ? '✓' : 'ℹ'}
                </div>
                <div>
                  <div style={{ fontSize: 13, color: '#1e293b', fontWeight: notif.read ? 500 : 700 }}>{notif.message}</div>
                  <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{notif.timestamp}</div>
                </div>
              </div>

              <div>
                {badge(notif.priority === 'High' ? 'red' : notif.priority === 'Medium' ? 'orange' : 'blue', notif.priority)}
              </div>
            </div>
          ))}
        </div>
      </div>
    );
}
