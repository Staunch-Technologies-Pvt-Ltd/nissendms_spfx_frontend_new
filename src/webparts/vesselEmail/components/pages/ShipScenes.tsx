// Ship-journey animations — DESIGN ONLY (no state, API or routing changes).
//  • <HelmWheel>          — CSS ship's wheel used on the login card + scenes
//  • <SignInVoyage>       — plays on "Continue with SharePoint": captain at the
//                            helm → camera turns to the bridge. Calls onDone()
//                            (which runs the unchanged goToSignIn redirect).
//  • <ShipTransitions>    — in AppLayout: arrival "Dashboard room" door opens
//                            once per session, and bulkhead doors swing open
//                            over every module change (host.state.view).
// Plays even when the OS 'reduce motion' setting is on (short, user-triggered).
// Set localStorage 'vdmsReduceMotion' = '1' to skip it.

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { skySeaBackgroundImage } from '../skySeaBackground';
import { dashboardShipBannerImage } from '../dashboardShipBanner';
import { vdmsFont } from '../futuristicTheme';

const reduced = (): boolean => {
  try { return typeof window !== 'undefined' && window.localStorage.getItem('vdmsReduceMotion') === '1'; } catch { return false; }
};

export function HelmWheel({ size, animation }: { size: number; animation?: string }): React.ReactElement {
  return (
    <div style={{ position: 'relative', width: size, height: size, animation, filter: 'drop-shadow(0 10px 18px rgba(0,0,0,0.35))' }}>
      {[0, 45, 90, 135].map(d => (
        <span key={d} style={{ position: 'absolute', left: '50%', top: '4%', width: size * 0.06, height: '92%', marginLeft: -size * 0.03, borderRadius: size * 0.03, background: 'linear-gradient(90deg,#7a4a22,#d0914f,#7a4a22)', transform: `rotate(${d}deg)` }} />
      ))}
      <span style={{ position: 'absolute', inset: '15%', borderRadius: '50%', border: `${size * 0.075}px solid #b8783c`, boxShadow: 'inset 0 0 0 2px rgba(0,0,0,0.2)' }} />
      {[0, 1, 2, 3, 4, 5, 6, 7].map(i => {
        const a = (i * Math.PI) / 4;
        return <span key={i} style={{ position: 'absolute', left: `${50 + Math.cos(a) * 46}%`, top: `${50 + Math.sin(a) * 46}%`, width: size * 0.11, height: size * 0.11, marginLeft: -size * 0.055, marginTop: -size * 0.055, borderRadius: '50%', background: 'radial-gradient(circle at 35% 30%,#f3c78a,#a8672d 70%)' }} />;
      })}
      <span style={{ position: 'absolute', left: '50%', top: '50%', width: size * 0.24, height: size * 0.24, marginLeft: -size * 0.12, marginTop: -size * 0.12, borderRadius: '50%', background: 'radial-gradient(circle at 35% 30%,#ffe7a8,#c9922f 55%,#7a5314)' }} />
    </div>
  );
}

function Caption({ title, sub }: { title: string; sub: string }): React.ReactElement {
  return (
    <div style={{ position: 'absolute', left: '50%', top: 34, transform: 'translateX(-50%)', padding: '12px 22px', borderRadius: 999, background: 'rgba(5,24,42,0.82)', border: '1px solid rgba(140,210,240,0.3)', color: '#eaf6fd', display: 'flex', alignItems: 'center', gap: 12, whiteSpace: 'nowrap', zIndex: 5, animation: 'vdms-cap .5s both', fontFamily: vdmsFont.ui }}>
      <span className="vdms-pulse" style={{ width: 9, height: 9, borderRadius: '50%', background: '#43e0a4' }} />
      <span style={{ fontSize: 16, fontWeight: 700 }}>{title}</span>
      <span style={{ fontFamily: vdmsFont.mono, fontSize: 13, color: '#9ab4c6' }}>{sub}</span>
    </div>
  );
}

function DoorLeaf({ side, animation }: { side: 'left' | 'right'; animation: string }): React.ReactElement {
  return (
    <div style={{ position: 'absolute', top: 0, bottom: 0, [side]: 0, width: '50%', transformOrigin: `${side} center`, backfaceVisibility: 'hidden', animation, background: `linear-gradient(${side === 'left' ? '100deg' : '260deg'}, #3d5f7e 0%, #24435f 55%, #17314a 100%)`, boxShadow: 'inset 0 0 0 3px rgba(255,255,255,0.08), inset 0 0 40px rgba(0,0,0,0.35)' }}>
      <div style={{ position: 'absolute', inset: 14, borderRadius: 10, border: '2px solid rgba(160,200,230,0.18)', backgroundImage: 'radial-gradient(circle, rgba(200,225,245,0.55) 0 2px, transparent 2.5px)', backgroundSize: '34px 34px', opacity: 0.5 }} />
      <div style={{ position: 'absolute', top: '22%', left: '50%', width: 92, height: 92, marginLeft: -46, borderRadius: '50%', border: '10px solid #8fa9c0', backgroundImage: `url("${skySeaBackgroundImage}")`, backgroundSize: '400%', backgroundPosition: '60% 40%' }} />
    </div>
  );
}

/** Login: helm → bridge, then onDone() (the unchanged SharePoint redirect). */
export function SignInVoyage({ onDone }: { onDone: () => void }): React.ReactElement | null {
  const [phase, setPhase] = React.useState<'helm' | 'bridge'>('helm');
  React.useEffect(() => {
    if (reduced()) { onDone(); return undefined; }
    const t1 = setTimeout(() => setPhase('bridge'), 1200);
    const t2 = setTimeout(onDone, 2900);
    return () => { clearTimeout(t1); clearTimeout(t2); };
  }, []);
  const w = typeof window !== 'undefined' ? window.innerWidth : 1200;
  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 100005, overflow: 'hidden', perspective: 1400, background: '#031423' }}>
      <div style={{ position: 'absolute', inset: 0, overflow: 'hidden', transformOrigin: 'right center', backfaceVisibility: 'hidden', animation: phase === 'bridge' ? 'vdms-cube-out 1s cubic-bezier(.6,0,.3,1) both' : 'vdms-fade .4s both' }}>
        <div style={{ position: 'absolute', inset: 0, backgroundImage: `url("${dashboardShipBannerImage}")`, backgroundSize: 'cover', backgroundPosition: 'center', animation: 'vdms-sail 5s ease-in-out infinite' }} />
        <Caption title="Captain at the helm" sub="setting course · Vessel DMS" />
        <div style={{ position: 'absolute', left: '50%', bottom: -Math.min(170, w * 0.18), marginLeft: -Math.min(210, w * 0.24) }}>
          <HelmWheel size={Math.min(420, w * 0.48)} animation="vdms-helm-spin 1.2s cubic-bezier(.5,0,.3,1) both" />
        </div>
      </div>
      {phase === 'bridge' && (
        <div style={{ position: 'absolute', inset: 0, overflow: 'hidden', background: 'linear-gradient(180deg,#0b2238,#061729)', transformOrigin: 'left center', backfaceVisibility: 'hidden', animation: 'vdms-cube-in 1s cubic-bezier(.6,0,.3,1) both' }}>
          <div style={{ position: 'absolute', left: '4%', right: '4%', top: '8%', height: '52%', display: 'flex', gap: 14 }}>
            {[0, 1, 2].map(i => <div key={i} style={{ flex: 1, borderRadius: '22px 22px 6px 6px', border: '12px solid #0b2238', borderBottomWidth: 18, backgroundImage: `url("${skySeaBackgroundImage}")`, backgroundSize: '300% auto', backgroundPosition: `${i * 50}% 45%` }} />)}
          </div>
          <div style={{ position: 'absolute', left: 0, right: 0, bottom: 0, height: '38%', background: 'linear-gradient(180deg,#123a5c,#0a2640)', borderTop: '3px solid rgba(52,213,234,0.45)' }}>
            <div style={{ position: 'absolute', right: '6%', top: '20%', padding: '14px 18px', borderRadius: 14, background: 'rgba(3,20,35,0.7)', border: '1px solid rgba(52,213,234,0.35)', color: '#7fe9f5', fontFamily: vdmsFont.mono, fontSize: 15, lineHeight: 1.7 }}>HDG 072°<br />SPD 14.2 kn<br />DEST · DASHBOARD</div>
          </div>
          <div style={{ position: 'absolute', left: '50%', bottom: -Math.min(120, w * 0.12), marginLeft: -Math.min(180, w * 0.22) }}>
            <HelmWheel size={Math.min(360, w * 0.44)} animation="vdms-helm-sway 1.6s ease-in-out infinite" />
          </div>
          <Caption title="On the bridge" sub="plotting course to Dashboard" />
        </div>
      )}
    </div>
  );
}

const LABELS: Record<string, [string, string]> = {
  dashboard: ['Dashboard', 'Home'], list: ['Documents', 'Documentation'], sites: ['Sites', 'SharepointLogo'], vessels: ['Vessels', 'Ferry'],
  migration: ['Migration Assistant', 'MoveToFolder'], users: ['User Management', 'People'], settings: ['Settings', 'Settings'],
  bento_email: ['AI Bento Email', 'Mail'], recycle: ['Recycle Bin', 'RecycleBin'], archive: ['Archive', 'Archive'], profile: ['Profile', 'Contact'],
};

/** Mount once inside AppLayout's root (next to <DeletionToastLayer />). */
export function ShipTransitions({ host }: { host: VesselEmail }): React.ReactElement | null {
  const view = String(host.state.view || 'dashboard');
  const [arrival, setArrival] = React.useState<boolean>(() => {
    try { return !reduced() && !sessionStorage.getItem('vdms-arrived'); } catch { return false; }
  });
  const [deck, setDeck] = React.useState<{ v: string; key: number } | null>(null);
  const prev = React.useRef(view);

  React.useEffect(() => {
    if (!arrival) return undefined;
    try { sessionStorage.setItem('vdms-arrived', '1'); } catch { /* ignore */ }
    const t = setTimeout(() => setArrival(false), 1500);
    return () => clearTimeout(t);
  }, []);

  React.useEffect(() => {
    if (prev.current === view) return undefined;
    prev.current = view;
    if (reduced()) return undefined;
    const key = Date.now();
    setDeck({ v: view, key });
    const t = setTimeout(() => setDeck(d => (d && d.key === key ? null : d)), 800);
    return () => clearTimeout(t);
  }, [view]);

  const vh = typeof window !== 'undefined' ? window.innerHeight : 800;
  const vw = typeof window !== 'undefined' ? window.innerWidth : 1200;

  if (arrival) {
    const dw = Math.min(560, vw * 0.72), dh = Math.min(720, vh * 0.74);
    return (
      <div style={{ position: 'fixed', inset: 0, zIndex: 100004, pointerEvents: 'none', animation: 'vdms-walk 1.5s cubic-bezier(.5,0,.3,1) both' }}>
        <div style={{ position: 'absolute', left: '50%', top: '53%', width: dw, height: dh, marginLeft: -dw / 2, marginTop: -dh / 2, borderRadius: '30px 30px 10px 10px', boxShadow: '0 0 0 18px #7f98ae, 0 0 0 22px #48647d, 0 0 0 200vmax #0a2238', perspective: 1400 }}>
          <DoorLeaf side="left" animation="vdms-leaf-l 1.5s cubic-bezier(.5,0,.3,1) both" />
          <DoorLeaf side="right" animation="vdms-leaf-r 1.5s cubic-bezier(.5,0,.3,1) both" />
          <div style={{ position: 'absolute', left: '50%', top: -74, transform: 'translateX(-50%)', padding: '10px 22px', borderRadius: 12, background: 'linear-gradient(135deg,#34d5ea,#2a8fe0)', color: '#02202b', fontFamily: vdmsFont.display, fontSize: 16, fontWeight: 800, letterSpacing: '.08em', whiteSpace: 'nowrap' }}>DECK 1 · DASHBOARD</div>
        </div>
      </div>
    );
  }
  if (!deck) return null;
  const [label, icon] = LABELS[deck.v] || [deck.v.replace(/_/g, ' '), 'Page'];
  return (
    <div key={deck.key} style={{ position: 'fixed', top: 96, left: 0, right: 0, bottom: 0, zIndex: 100003, pointerEvents: 'none', perspective: 1600, overflow: 'hidden' }}>
      <DoorLeaf side="left" animation="vdms-deck-l .8s cubic-bezier(.5,0,.3,1) both" />
      <DoorLeaf side="right" animation="vdms-deck-r .8s cubic-bezier(.5,0,.3,1) both" />
      <div style={{ position: 'absolute', left: '50%', top: '50%', transform: 'translate(-50%,-50%)', display: 'flex', alignItems: 'center', gap: 14, padding: '14px 22px', borderRadius: 18, background: 'var(--clay-accent-gradient)', color: '#fff', whiteSpace: 'nowrap', animation: 'vdms-plate .8s ease both', fontFamily: vdmsFont.ui }}>
        <Icon iconName={icon} style={{ fontSize: 24 }} />
        <div>
          <div style={{ fontFamily: vdmsFont.mono, fontSize: 11, fontWeight: 600, letterSpacing: '.2em', textTransform: 'uppercase', opacity: 0.85 }}>Entering deck</div>
          <div style={{ fontFamily: vdmsFont.display, fontSize: 22, fontWeight: 800 }}>{label}</div>
        </div>
      </div>
    </div>
  );
}
