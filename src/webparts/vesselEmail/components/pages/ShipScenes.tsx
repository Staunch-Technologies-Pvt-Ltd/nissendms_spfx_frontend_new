// Sign-in transition — DESIGN ONLY (no state, API or routing changes).
//  • <SignInVoyage>    — plays briefly on "Continue with SharePoint" before
//                         calling onDone() (the unchanged goToSignIn redirect).
//  • <ShipTransitions> — mounted once inside AppLayout; intentionally renders
//                         nothing (the previous per-navigation door animation
//                         was removed as part of the enterprise UI redesign).

import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { vdmsFont } from '../futuristicTheme';

/** Login: a brief "redirecting" overlay, then onDone() (the unchanged SharePoint redirect). */
export function SignInVoyage({ onDone }: { onDone: () => void }): React.ReactElement {
  React.useEffect(() => {
    const t = setTimeout(onDone, 220);
    return () => clearTimeout(t);
  }, []);
  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 100005,
      background: '#0b2a4a', color: '#ffffff',
      display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 14,
      fontFamily: vdmsFont.ui,
    }}>
      <div style={{
        width: 28, height: 28, borderRadius: '50%',
        border: '3px solid rgba(255,255,255,0.25)', borderTopColor: '#2dd4bf',
        animation: 'vdms-spin 0.8s linear infinite',
      }} />
      <div style={{ fontSize: 15, fontWeight: 600, letterSpacing: '0.02em' }}>Redirecting to sign-in…</div>
      <style>{'@keyframes vdms-spin { to { transform: rotate(360deg) } }'}</style>
    </div>
  );
}

/** Mounted once inside AppLayout's root. Reserved for future transition
 * treatments — deliberately a no-op today. */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
export function ShipTransitions({ host }: { host: VesselEmail }): React.ReactElement | null {
  return null;
}
