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

export function renderLayout(host: VesselEmail, content: React.ReactElement): React.ReactElement {
    const userDisplayName = host.props.userDisplayName || 'Priya';
    return (
      <div style={{ display: 'flex', alignItems: 'stretch', minHeight: '100vh', background: '#f8fafc', width: '100%', fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" }}>
        {host._renderSidebar()}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 }}>
          {/* Top Bar Header */}
          <div style={{
            height: 52, background: '#ffffff', borderBottom: '1px solid #e2e8f0',
            display: 'flex', alignItems: 'center', justifyContent: 'space-between',
            padding: '0 24px', flexShrink: 0,
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ color: '#64748b', fontSize: 14, fontWeight: 500 }}>Vessel DMS</span>
              <span style={{ color: '#cbd5e1' }}>/</span>
              <span style={{ color: '#0f172a', fontSize: 14, fontWeight: 600, textTransform: 'capitalize' }}>
                {host.state.view.replace('_', ' ')}
              </span>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
              <div style={{ position: 'relative', width: 220 }}>
                <input
                  type="text"
                  placeholder="Search host site..."
                  style={{
                    width: '100%', padding: '6px 12px 6px 30px', borderRadius: 16,
                    border: '1px solid #cbd5e1', fontSize: 12, background: '#f8fafc',
                    outline: 'none', boxSizing: 'border-box',
                  }}
                />
                <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 12 }}>🔍</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{
                  width: 28, height: 28, borderRadius: '50%', background: '#0078d4',
                  color: '#fff', fontSize: 12, fontWeight: 600, display: 'flex',
                  alignItems: 'center', justifyContent: 'center',
                }}>
                  {userDisplayName.charAt(0)}
                </div>
                <span style={{ fontSize: 13, fontWeight: 600, color: '#334155' }}>{userDisplayName}</span>
              </div>
            </div>
          </div>

          {/* Main Module Content */}
          <div style={{ flex: 1, padding: 24, overflowY: 'auto' }}>
            {content}
          </div>

          {/* Right-side Aside Document Preview Drawer */}
          {host._renderDocPreviewDrawer()}
        </div>
      </div>
    );
}
