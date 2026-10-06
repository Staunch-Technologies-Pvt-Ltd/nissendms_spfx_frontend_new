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
import { Icon } from '@fluentui/react/lib/Icon';

export function renderDocPreviewDrawer(host: VesselEmail): React.ReactElement | null {
    const { selectedDocPreview } = host.state;
    if (!selectedDocPreview) return null;

    const base = host._base();
    const downloadUrl = selectedDocPreview.fileId ? `${base}/api/files/${selectedDocPreview.fileId}/content` : null;

    return (
      <div
        style={{
          position: 'fixed',
          top: 0,
          right: 0,
          bottom: 0,
          width: 'min(380px, 100vw)',
          maxWidth: '90vw',
          background: '#ffffff',
          boxShadow: '-6px 0 24px rgba(0, 0, 0, 0.18)',
          zIndex: 10000,
          display: 'flex',
          flexDirection: 'column',
          borderLeft: '1px solid #cbd5e1',
          animation: 'slideIn 0.2s ease-out',
        }}
      >
        {/* Drawer Header */}
        <div style={{
          padding: '16px 20px', background: '#0f172a', color: '#fff',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontWeight: 700, fontSize: 15 }}>
            <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 15 }} /> Document Overview
          </div>
          <button
            onClick={() => host.setState({ selectedDocPreview: null })}
            style={{ border: 'none', background: 'transparent', color: '#cbd5e1', fontSize: 18, cursor: 'pointer', fontWeight: 700 }}
            title="Close Drawer"
          >
            <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 18 }} />
          </button>
        </div>

        {/* Drawer Content */}
        <div style={{ flex: 1, padding: 20, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 }}>
          {/* File Icon Banner */}
          <div style={{
            background: '#f0f9ff', borderRadius: 12, padding: 20, border: '1px solid #bae6fd',
            display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', gap: 8,
          }}>
            <div style={{ fontSize: 42 }}><Icon iconName="Page" aria-hidden="true" style={{ fontSize: 42 }} /></div>
            <a
              href={downloadUrl || '#'}
              target="_blank"
              rel="noopener noreferrer"
              onClick={e => {
                if (!downloadUrl) {
                  e.preventDefault();
                  alert(`File "${selectedDocPreview.fileName}" stored locally.`);
                }
              }}
              style={{
                fontSize: 15, fontWeight: 700, color: '#0284c7', textDecoration: 'underline',
                wordBreak: 'break-all', cursor: 'pointer',
              }}
            >
              {selectedDocPreview.fileName}
            </a>
            <span style={{ fontSize: 11, background: '#dbeafe', color: '#1e40af', padding: '2px 8px', borderRadius: 10, fontWeight: 600 }}>
              Uploaded Document
            </span>
          </div>

          {/* Document Properties */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 14, display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 2 }}>
              Metadata Details
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}><Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /> Vessel Name:</span>
              <strong style={{ color: '#0f172a' }}>{selectedDocPreview.vesselName}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}><Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 12 }} /> Group:</span>
              <span style={{ color: '#2563eb', fontWeight: 600 }}>{selectedDocPreview.group}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}><Icon iconName="Tag" aria-hidden="true" style={{ fontSize: 12 }} /> Category:</span>
              <strong style={{ color: '#1e293b' }}>{selectedDocPreview.category}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}><Icon iconName="Save" aria-hidden="true" style={{ fontSize: 12 }} /> Size:</span>
              <span style={{ color: '#475569' }}>{selectedDocPreview.size || '142 KB'}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}><Icon iconName="Calendar" aria-hidden="true" style={{ fontSize: 12 }} /> Date:</span>
              <span style={{ color: '#475569' }}>{selectedDocPreview.date || 'Today'}</span>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2, fontSize: 12 }}>
              <span style={{ color: '#64748b' }}><Icon iconName="Globe2" aria-hidden="true" style={{ fontSize: 12 }} /> Folder Path:</span>
              <span style={{ color: '#334155', fontFamily: 'monospace', fontSize: 11, background: '#f8fafc', padding: '4px 6px', borderRadius: 4, wordBreak: 'break-all' }}>
                {selectedDocPreview.folderPath}
              </span>
            </div>
          </div>

          {/* Quick Preview Card */}
          <div style={{ background: '#f8fafc', borderRadius: 10, border: '1px dashed #cbd5e1', padding: 14, textAlign: 'center' }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>
              Document Status: <span style={{ color: '#16a34a', fontWeight: 700 }}>Active / Ready</span>
            </div>
            <p style={{ margin: 0, fontSize: 11, color: '#64748b' }}>
              Click open below to view or download the uploaded document directly.
            </p>
          </div>
        </div>

        {/* Drawer Actions Footer */}
        <div style={{ padding: 16, background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <a
            href={downloadUrl || '#'}
            target="_blank"
            rel="noopener noreferrer"
            onClick={e => {
              if (!downloadUrl) {
                e.preventDefault();
                alert(`File "${selectedDocPreview.fileName}" stored locally.`);
              }
            }}
            style={{
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
              background: '#0284c7', color: '#fff', padding: '9px 16px', borderRadius: 8,
              fontSize: 13, fontWeight: 600, textDecoration: 'none', textAlign: 'center',
            }}
          >
            <Icon iconName="Globe" aria-hidden="true" style={{ fontSize: 13 }} /> Open / View Document
          </a>
          <button
            onClick={() => {
              host.setState({
                view: 'bento_email',
                bentoUploadVessel: selectedDocPreview.vesselName,
                selectedDocPreview: null,
              });
            }}
            style={{
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
              background: '#0f172a', color: '#fff', border: 'none', padding: '9px 16px', borderRadius: 8,
              fontSize: 13, fontWeight: 600, cursor: 'pointer',
            }}
          >
            <Icon iconName="Mail" aria-hidden="true" style={{ fontSize: 13 }} /> Send via Bento Email
          </button>
          <button
            onClick={() => host.setState({ selectedDocPreview: null })}
            style={{
              background: '#fff', border: '1px solid #cbd5e1', color: '#475569', padding: '8px 16px',
              borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            Close Drawer
          </button>
        </div>
      </div>
    );
}
