import * as React from 'react';
import * as ReactDOM from 'react-dom';
import type VesselEmail from '../VesselEmail';
import { isMobileWidth } from '../responsive';
import { Icon } from '@fluentui/react/lib/Icon';

/** Opened from the Documents module's "New Folder" toolbar button (top,
 *  next to Archive). `folderRef` is whatever the backend's
 *  `_resolve_drive_folder_id` accepts: 'root', a Graph item id, or a
 *  slash/" > "-delimited path — the same shape `displayPath` shows the
 *  user, so leaving the path field untouched reuses it verbatim. */
export type CreateFolderDialogState = {
  siteId: string;
  driveId: string;
  folderRef: string;
  displayPath: string;
  vesselName: string;
};

export function renderCreateFolderModal(host: VesselEmail): React.ReactElement | null {
  const dialog = host.state.createFolderDialog;
  if (!dialog) return null;
  return <CreateFolderDialog key={`${dialog.siteId}::${dialog.driveId}::${dialog.folderRef}`} host={host} dialog={dialog} />;
}

function CreateFolderDialog({ host, dialog }: { host: VesselEmail; dialog: CreateFolderDialogState }): React.ReactElement {
  const { siteId, driveId, folderRef, displayPath } = dialog;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  const [path, setPath] = React.useState(displayPath || '');
  const [name, setName] = React.useState('');
  const [asVessel, setAsVessel] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [resultMsg, setResultMsg] = React.useState<string | null>(null);

  const close = (): void => {
    if (busy) return;
    host.setState({ createFolderDialog: null });
  };

  const submit = async (): Promise<void> => {
    const trimmedName = name.trim();
    if (!trimmedName) { setError('Enter a folder name.'); return; }

    if (asVessel) {
      // The full vessel record (IMO, hull number, site/parent-folder picker,
      // etc.) already has a dedicated, validated flow — reopen it instead of
      // re-implementing vessel creation here, just pre-filling the name the
      // user already typed.
      host.setState({ createFolderDialog: null });
      host._openCreate(trimmedName);
      return;
    }

    setBusy(true);
    setError(null);
    const trimmedPath = path.trim();
    // Only override the location this popup was opened at if the user
    // actually edited the path field.
    const targetRef = trimmedPath && trimmedPath !== (displayPath || '').trim()
      ? trimmedPath
      : folderRef;
    const result = await host._createFolderAtPath(siteId, driveId, targetRef, trimmedName);
    setBusy(false);
    if (!result.success) {
      setError(result.error || 'Could not create the folder.');
      return;
    }
    setResultMsg(`"${trimmedName}" created successfully.`);
    setTimeout(() => host.setState({ createFolderDialog: null }), 1200);
  };

  const themeMode = host.state.themeMode === 'night' ? 'night' : 'light';

  // Rendered via a portal onto <body> for the same reason ArchiveSelectionModal
  // is — the SharePoint workbench's ancestor `transform` otherwise clips a
  // `position: fixed` dialog to the page canvas instead of the real viewport.
  return ReactDOM.createPortal(
    <div data-vessel-theme={themeMode} style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 100001,
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 20,
    }} onClick={close}>
      <div
        style={{
          background: 'var(--vdms-surface)', borderRadius: 16, padding: isMobile ? '16px 14px' : '24px 28px',
          width: isMobile ? '95vw' : 480, maxWidth: '95vw', display: 'flex', flexDirection: 'column',
          boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
        }}
        onClick={e => e.stopPropagation()}
      >
        {resultMsg ? (
          <div style={{ textAlign: 'center', padding: '20px 0 4px' }}>
            <div style={{ fontSize: 44, marginBottom: 10 }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 44 }} /></div>
            <div style={{ fontSize: 17, fontWeight: 700, color: '#059669', marginBottom: 6 }}>Folder created</div>
            <p style={{ fontSize: 13, color: 'var(--vdms-text-muted)', margin: 0 }}>{resultMsg}</p>
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
              <div style={{ width: 40, height: 40, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}><Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 20 }} /></div>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--vdms-text)' }}>New folder</div>
                <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                  Creates a folder in SharePoint at the path below
                </div>
              </div>
            </div>

            <label style={{ fontSize: 12, fontWeight: 700, color: 'var(--vdms-text-muted)', marginBottom: 4, textTransform: 'uppercase', letterSpacing: '0.03em' }}>
              Folder name
            </label>
            <input
              type="text"
              value={name}
              onChange={e => setName(e.target.value)}
              placeholder="e.g. 2026 Survey Reports"
              disabled={busy}
              autoFocus
              style={{
                width: '100%', boxSizing: 'border-box', borderRadius: 8, border: '1px solid var(--vdms-border)',
                padding: '9px 12px', fontSize: 13, marginBottom: 14, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)',
              }}
            />

            {!asVessel && (
              <>
                <label style={{ fontSize: 12, fontWeight: 700, color: 'var(--vdms-text-muted)', marginBottom: 4, textTransform: 'uppercase', letterSpacing: '0.03em' }}>
                  Path to create it in
                </label>
                <input
                  type="text"
                  value={path}
                  onChange={e => setPath(e.target.value)}
                  disabled={busy}
                  placeholder="Documents"
                  style={{
                    width: '100%', boxSizing: 'border-box', borderRadius: 8, border: '1px solid var(--vdms-border)',
                    padding: '9px 12px', fontSize: 12.5, marginBottom: 6, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)',
                    fontFamily: 'monospace',
                  }}
                />
                <div style={{ fontSize: 11, color: 'var(--vdms-text-faint)', marginBottom: 14 }}>
                  Defaults to where you opened this from — edit it to create the folder somewhere else instead.
                </div>
              </>
            )}

            <label style={{
              display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', borderRadius: 8,
              background: asVessel ? '#eff6ff' : 'var(--vdms-surface-alt)', border: `1px solid ${asVessel ? '#bfdbfe' : 'var(--vdms-border)'}`,
              cursor: busy ? 'default' : 'pointer', marginBottom: 16,
            }}>
              <input
                type="checkbox"
                checked={asVessel}
                disabled={busy}
                onChange={e => setAsVessel(e.target.checked)}
                style={{ width: 15, height: 15, accentColor: '#2563eb', cursor: 'pointer' }}
              />
              <span style={{ fontSize: 12.5, color: 'var(--vdms-text)' }}>
<Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12.5 }} /> This is a new vessel — open the Add Vessel form instead
              </span>
            </label>

            {error && (
              <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#dc2626', marginBottom: 14 }}>
<Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {error}
              </div>
            )}

            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', flexDirection: isMobile ? 'column' : 'row' }}>
              <button
                onClick={close}
                disabled={busy}
                style={{ minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface)', fontSize: 13, fontWeight: 600, cursor: 'pointer', color: 'var(--vdms-text)' }}
              >
                Cancel
              </button>
              <button
                onClick={() => void submit()}
                disabled={busy || !name.trim()}
                style={{
                  minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 600,
                  cursor: busy || !name.trim() ? 'not-allowed' : 'pointer',
                  background: busy || !name.trim() ? '#93c5fd' : '#2563eb', color: '#fff',
                }}
              >
                {busy ? 'Creating…' : asVessel ? 'Continue to Add Vessel' : 'Create folder'}
              </button>
            </div>
          </>
        )}
      </div>
    </div>,
    document.body
  );
}
