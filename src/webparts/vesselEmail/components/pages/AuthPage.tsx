import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { vdmsFont, injectFuturisticTheme } from '../futuristicTheme';
import { SignInVoyage } from './ShipScenes';

export type AuthPageMode = 'login' | 'logout';

// UNCHANGED — design-only update. Sign in and Sign up both use the same
// Microsoft 365 (Entra ID) flow; "Sign up" is a visual tab only.
function goToSignIn(host: VesselEmail): void {
  const returnUrl = window.location.href;
  const siteUrl = (host.props.siteUrl || window.location.origin).replace(/\/$/, '');
  window.location.assign(`${siteUrl}/_layouts/15/authenticate.aspx?Source=${encodeURIComponent(returnUrl)}`);
}

function tenantHost(host: VesselEmail): string {
  try { return new URL(host.props.siteUrl || window.location.origin).hostname; } catch { return ''; }
}

export function renderAuthPage(host: VesselEmail, mode: AuthPageMode): React.ReactElement {
  return <AuthScreen host={host} mode={mode} />;
}

function AuthScreen({ host, mode }: { host: VesselEmail; mode: AuthPageMode }): React.ReactElement {
  injectFuturisticTheme();
  const [voyage, setVoyage] = React.useState(false);
  const [tab, setTab] = React.useState<'signin' | 'signup'>('signin');

  const isLogout = mode === 'logout';
  const signUp = tab === 'signup' && !isLogout;
  const displayName = host.props.userDisplayName || '';
  const tenant = tenantHost(host) || 'contoso.sharepoint.com';

  const blurb = isLogout ? 'Your Vessel DMS session has been closed. Return to SharePoint to start a new session.'
    : signUp ? "Accounts use your organisation's Microsoft 365 identity. We'll send you to Microsoft to finish."
    : `Continue with your SharePoint account${displayName ? `, ${displayName}` : ''} to access vessel documents and workflows.`;
  const cta = isLogout ? 'Return to sign in' : signUp ? 'Sign up with SharePoint' : 'Continue with SharePoint';

  const tabBtn = (t: 'signin' | 'signup'): React.CSSProperties => ({
    border: 'none', cursor: 'pointer', fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 700, padding: '10px 0',
    borderRadius: 8, background: tab === t ? '#0b2a4a' : 'transparent', color: tab === t ? '#fff' : '#334155',
  });

  return (
    <div
      className="vessel-dms-auth"
      data-vessel-theme={host.state.themeMode}
      style={{
        position: 'fixed', inset: 0, zIndex: 99999, overflow: 'auto', fontFamily: vdmsFont.ui,
        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, boxSizing: 'border-box',
      }}
    >
      {voyage && <SignInVoyage onDone={() => goToSignIn(host)} />}

      <div className="vdms-pop" style={{
        boxSizing: 'border-box', width: 'min(440px, 100%)', padding: '40px 36px', borderRadius: 16,
        background: '#ffffff', boxShadow: '0 1px 2px rgba(0,0,0,0.08), 0 20px 48px rgba(0,0,0,0.28)',
        textAlign: 'center', color: '#101b2d',
      }}>
        <div style={{
          width: 56, height: 56, borderRadius: 12, margin: '0 auto 18px', background: '#0b2a4a',
          display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff',
        }}>
          <Icon iconName="Ferry" style={{ fontSize: 28 }} />
        </div>
        <h1 style={{ margin: 0, fontFamily: vdmsFont.display, fontSize: 30, fontWeight: 800, letterSpacing: '-0.02em' }}>
          Vessel DMS
        </h1>
        <div style={{ marginTop: 6, fontSize: 14, fontWeight: 500, color: '#5b6b7f' }}>
          Your Vessel Documents, Always Accessible
        </div>

        {!isLogout && (
          <div role="tablist" style={{ marginTop: 24, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 4, padding: 4, borderRadius: 10, background: '#f1f5f9' }}>
            <button type="button" role="tab" aria-selected={tab === 'signin'} onClick={() => setTab('signin')} style={tabBtn('signin')}>Sign in</button>
            <button type="button" role="tab" aria-selected={tab === 'signup'} onClick={() => setTab('signup')} style={tabBtn('signup')}>Sign up</button>
          </div>
        )}

        <p style={{ margin: '18px 0', fontSize: 14, lineHeight: 1.5, color: '#475569', fontWeight: 500 }}>{blurb}</p>

        <button
          type="button"
          onClick={() => setVoyage(true)}
          style={{
            width: '100%', minHeight: 48, border: 'none', borderRadius: 10, cursor: 'pointer',
            fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 700, color: '#fff', background: '#0e7490',
            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
          }}
        >
          <Icon iconName="Permissions" style={{ fontSize: 16 }} />
          <span>{cta}</span>
        </button>

        <div title="Signing in to this SharePoint tenant" style={{ marginTop: 16 }}>
          <div style={{
            boxSizing: 'border-box', padding: '10px 14px', borderRadius: 10, border: '1px solid #e2e8f0',
            background: '#f8fafc', fontSize: 13, fontWeight: 500, color: '#334155', textAlign: 'left',
            overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'flex', alignItems: 'center', gap: 8,
          }}>
            <Icon iconName="Globe" style={{ fontSize: 13, color: '#64748b' }} />
            {tenant}
          </div>
        </div>

        <div style={{ marginTop: 16, display: 'inline-flex', alignItems: 'center', gap: 8, padding: '6px 14px', borderRadius: 999, background: '#f0fdf4', fontFamily: vdmsFont.mono, fontSize: 12, color: '#15803d' }}>
          <span className="vdms-pulse" style={{ width: 7, height: 7, borderRadius: '50%', background: '#22c55e' }} />
          All services operational
        </div>

        <div style={{ marginTop: 16, fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6 }}>
          <Icon iconName="Lock" style={{ fontSize: 11 }} />
          Secure access is managed by your organization.
        </div>
      </div>
    </div>
  );
}
