import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';

export function renderProfilePage(host: VesselEmail): React.ReactElement {
  const displayName = host.props.userDisplayName || 'Admin';
  const email = host.props.userEmail || 'SharePoint account';
  const initial = displayName.charAt(0).toUpperCase();

  return (
    <div style={{ maxWidth: 860, margin: '0 auto', display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div>
        <h1 style={{ margin: 0, color: clay.text, fontSize: 28, fontWeight: 800 }}>Profile</h1>
        <p style={{ margin: '6px 0 0', color: clay.textMuted, fontSize: 15 }}>Your Vessel DMS account details and shortcuts.</p>
      </div>

      <section style={{ background: clay.surface, border: `1px solid ${clay.accentSoft}`, borderRadius: 20, padding: 28, boxShadow: clay.shadowRaised, display: 'flex', alignItems: 'center', gap: 20 }}>
        <div style={{ width: 76, height: 76, borderRadius: 24, background: clay.accentGradient, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 32, fontWeight: 900, boxShadow: clay.shadowButton }}>
          {initial}
        </div>
        <div>
          <h2 style={{ margin: 0, color: clay.text, fontSize: 22, fontWeight: 800 }}>{displayName}</h2>
          <div style={{ marginTop: 6, color: clay.textMuted, fontSize: 15 }}>{email}</div>
          <div style={{ marginTop: 10, display: 'inline-flex', padding: '4px 10px', borderRadius: 999, background: clay.accentSoft, color: clay.accentDark, fontSize: 12, fontWeight: 800 }}>Vessel DMS user</div>
        </div>
      </section>

      <section style={{ background: clay.surface, border: `1px solid ${clay.accentSoft}`, borderRadius: 20, padding: 24, boxShadow: clay.shadowRaised }}>
        <h2 style={{ margin: '0 0 6px', color: clay.text, fontSize: 18, fontWeight: 800 }}>Account shortcuts</h2>
        <p style={{ margin: '0 0 18px', color: clay.textMuted, fontSize: 14 }}>Manage application preferences and your active session.</p>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
          <button type="button" onClick={() => host._goToView('settings')} style={{ background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: 12, padding: '11px 18px', fontWeight: 800, cursor: 'pointer', boxShadow: clay.shadowButton }}>Open Settings</button>
          <button type="button" onClick={host._handleSignOut} style={{ background: clay.surfaceRaised, color: clay.accentDeep, border: `1px solid ${clay.accentSoft}`, borderRadius: 12, padding: '11px 18px', fontWeight: 800, cursor: 'pointer' }}>Sign out</button>
        </div>
      </section>
    </div>
  );
}