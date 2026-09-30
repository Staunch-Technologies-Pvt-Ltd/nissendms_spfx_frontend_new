import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';

export type AuthPageMode = 'login' | 'logout';

// Send the user to SharePoint's own sign-in page (Entra ID). If a valid
// Microsoft session already exists it returns immediately with a fresh token;
// otherwise the user sees the login screen. Source brings them back to the
// exact page (including workbench debug params) they came from.
function goToSignIn(host: VesselEmail): void {
  const returnUrl = window.location.href;
  const siteUrl = (host.props.siteUrl || window.location.origin).replace(/\/$/, '');
  window.location.assign(`${siteUrl}/_layouts/15/authenticate.aspx?Source=${encodeURIComponent(returnUrl)}`);
}

export function renderAuthPage(host: VesselEmail, mode: AuthPageMode): React.ReactElement {
  const isLogout = mode === 'logout';
  const displayName = host.props.userDisplayName || 'SharePoint user';
  const isNight = host.state.themeMode === 'night';

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, background: isNight ? '#211812' : clay.bg, fontFamily: "'Segoe UI Variable', 'Segoe UI', sans-serif" }}>
      <div style={{ width: '100%', maxWidth: 470, background: isNight ? '#302219' : clay.surface, border: `1px solid ${isNight ? '#614331' : clay.accentSoft}`, borderRadius: 26, padding: '44px 42px', textAlign: 'center', boxShadow: clay.shadowRaised }}>
        <div style={{ width: 72, height: 72, margin: '0 auto 20px', borderRadius: 22, background: clay.accentGradient, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, fontWeight: 900, boxShadow: clay.shadowButton }}>VD</div>
        <div style={{ color: clay.accentDark, fontSize: 12, fontWeight: 800, letterSpacing: '0.14em', textTransform: 'uppercase' }}>Vessel Documents</div>
        <h1 style={{ margin: '12px 0 10px', color: isNight ? '#f8eee6' : clay.text, fontSize: 28, fontWeight: 850 }}>{isLogout ? 'You are signed out' : 'Sign in to Vessel DMS'}</h1>
        <p style={{ margin: '0 auto 26px', maxWidth: 360, color: isNight ? '#c7a58d' : clay.textMuted, fontSize: 15, lineHeight: 1.6 }}>
          {isLogout ? 'Your Vessel DMS session has been closed. Return to SharePoint to start a new session.' : `Continue with your SharePoint account${displayName ? `, ${displayName}` : ''} to access vessel documents and workflows.`}
        </p>
        <button
          type="button"
          onClick={() => goToSignIn(host)}
          style={{ width: '100%', minHeight: 48, border: 'none', borderRadius: 14, background: clay.accentGradient, color: '#fff', fontSize: 15, fontWeight: 800, cursor: 'pointer', boxShadow: clay.shadowButton, transition: 'filter 0.16s ease, transform 0.16s ease' }}
          onMouseEnter={e => { e.currentTarget.style.filter = 'brightness(0.88) saturate(1.12)'; e.currentTarget.style.transform = 'translateY(-2px)'; }}
          onMouseLeave={e => { e.currentTarget.style.filter = 'none'; e.currentTarget.style.transform = 'translateY(0)'; }}
        >
          <Icon iconName={isLogout ? 'Refresh' : 'Contact'} style={{ marginRight: 8 }} />
          {isLogout ? 'Return to sign in' : 'Continue with SharePoint'}
        </button>
        <div style={{ marginTop: 18, color: isNight ? '#a8836c' : clay.textMuted, fontSize: 12 }}>Secure access is managed by your organization.</div>
      </div>
    </div>
  );
}