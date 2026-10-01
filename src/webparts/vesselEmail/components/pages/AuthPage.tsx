
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { vdmsFont, injectFuturisticTheme } from '../futuristicTheme';
import { loginSceneImage, LOGIN_SCENE_W as SW, LOGIN_SCENE_H as SH, LOGIN_SHIP_POLY } from '../loginScene';
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

const reduceMotion = (): boolean => {
  try { return window.localStorage.getItem('vdmsReduceMotion') === '1'; } catch { return false; }
};

/** Animated sea + drifting clouds + sun drawn over the scene photo. The ship area is never displaced. */
function useSceneAnimation(ref: React.RefObject<HTMLCanvasElement>, night: boolean, lRef: React.RefObject<HTMLCanvasElement>, rRef: React.RefObject<HTMLCanvasElement>, mRef: React.MutableRefObject<number>): void {
  React.useEffect(() => {
    const cv = ref.current; if (!cv) return undefined;
    const g = cv.getContext('2d'); if (!g) return undefined;
    const img = new Image(); img.src = loginSceneImage;
    const poly = LOGIN_SHIP_POLY;
    const inShip = (x: number, y: number): boolean => {
      let c = false;
      for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
        const a = poly[i], b = poly[j];
        if ((a[1] > y) !== (b[1] > y) && x < (b[0] - a[0]) * (y - a[1]) / (b[1] - a[1]) + a[0]) c = !c;
      }
      return c;
    };
    const rn = (n: number): number => { const v = Math.sin(n * 91.7) * 43758.5; return v - Math.floor(v); };
    const still = reduceMotion();
    let hi: HTMLCanvasElement | null = null;   // 2x-resolution, slightly enriched copy of the photo
    let cloud: HTMLCanvasElement | null = null; // clouds-only layer (alpha from brightness) so the sun can pass behind
    let ext: Array<[number, number] | null> = [];
    const flecks: Array<{ x: number; y: number; v: number; p: number; w: number }> = [];
    for (let i = 0; i < 260; i++) {
      const x = rn(i) * SW, y = 250 + rn(i + 500) * (SH - 250);
      if (!inShip(x, y) && !(x > 1000 && y > 80 && y < 840 && x < 1620)) flecks.push({ x, y, v: 0.3 + rn(i + 9) * 0.9, p: rn(i + 3) * 6, w: 6 + rn(i + 7) * 26 * (0.4 + (y - 250) / (SH - 250)) });
    }
    const prep = (): void => {
      hi = document.createElement('canvas'); hi.width = SW * 2; hi.height = SH * 2;
      const h = hi.getContext('2d') as CanvasRenderingContext2D;
      h.imageSmoothingQuality = 'high'; h.filter = 'contrast(1.07) saturate(1.15) brightness(1.03)';
      h.drawImage(img, 0, 0, SW * 2, SH * 2);
      cloud = document.createElement('canvas'); cloud.width = SW * 2; cloud.height = 524;
      const k = cloud.getContext('2d') as CanvasRenderingContext2D;
      k.drawImage(hi, 0, 0, SW * 2, 524, 0, 0, SW * 2, 524);
      const id = k.getImageData(0, 0, SW * 2, 524), p = id.data;
      for (let i = 0; i < p.length; i += 4) p[i + 3] = Math.max(0, Math.min(255, (Math.min(p[i], p[i + 1], p[i + 2]) - 135) * 4.2));
      k.putImageData(id, 0, 0);
      ext = [];
      for (let y = 0; y < SH; y += 3) {
        let lo = 9999, hiX = -1;
        for (let x = 100; x < 1100; x += 4) if (inShip(x, y + 1)) { lo = Math.min(lo, x); hiX = Math.max(hiX, x); }
        ext[y / 3] = hiX < 0 ? null : [lo - 22, hiX + 22];
      }
    };
    let timer = 0;
    const tick = (): void => {
      if (!hi) { if (img.complete && img.naturalWidth) prep(); else return; }
      const T = still ? 0 : Date.now() / 1000;
      g.setTransform(2, 0, 0, 2, 0, 0); g.clearRect(0, 0, SW, SH);
      const d = (sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void =>
        g.drawImage(hi as HTMLCanvasElement, sx * 2, sy * 2, sw * 2, sh * 2, dx, dy, dw, dh);
      d(0, 0, SW, SH, 0, 0, SW, SH);
      const row = (y: number, dx: number, dy: number, x0: number, x1: number): void => { if (x1 > x0) d(x0 - dx, y, x1 - x0, 3, x0, y + dy, x1 - x0, 3); };
      for (let y = 0; y < 250; y += 3) d(0, y, SW, 3, Math.sin(T * 0.22 + y * 0.01) * (6 + (250 - y) * 0.06), y, SW, 3);
      for (let y = 246; y < SH; y += 3) {
        const f = (y - 246) / (SH - 246), am = 1.5 + f * 7, e = ext[y / 3];
        const dx = am * (Math.sin(y * 0.05 - T * 1.7) + 0.5 * Math.sin(y * 0.13 + T * 2.4)), dy = Math.sin(T * 1.3 + y * 0.02) * f * 1.2;
        if (e) { row(y, dx, dy, 0, e[0]); row(y, dx, dy, e[1], SW); } else row(y, dx, dy, 0, SW);
      }
      d(0, 735, 830, 184, 0, 735, 830, 184); // feature row stays still (keeps its text crisp)
      // Side margins: mirrored copies of the (animated) scene edges, so the page fills any screen width with no bars
      const m = Math.min(SW, Math.ceil(mRef.current)) + 2, m2 = m * 2, lc = lRef.current, rc = rRef.current;
      const side = (c: HTMLCanvasElement | null, sx: number): CanvasRenderingContext2D | null => {
        const x = c && c.getContext('2d'); if (!x || !c) return null;
        x.setTransform(1, 0, 0, 1, 0, 0); x.clearRect(0, 0, c.width, c.height);
        if (mRef.current > 0) { x.setTransform(-1, 0, 0, 1, SW * 2, 0); x.drawImage(cv, sx, 0, m2, SH * 2, sx, 0, m2, SH * 2); }
        x.setTransform(2, 0, 0, 2, 0, 0);
        return x;
      };
      const gl = side(lc, 0), gr = side(rc, SW * 2 - m2);
      // Sun: small yellow disc, travels left → right across the whole screen width, passes behind the clouds and the ship
      const mm = mRef.current, ph = T ? ((T / 70) % 1) : 0.3, SX = -mm - 20 + ph * (SW + 2 * mm + 40), SY = 235 - Math.sin(Math.PI * ph) * 150;
      const disc = (c: CanvasRenderingContext2D, cx: number): void => {
        c.save(); c.globalCompositeOperation = 'screen';
        // soft warm halo
        const halo = c.createRadialGradient(cx, SY, 6, cx, SY, 110);
        halo.addColorStop(0, 'rgba(255,236,170,0.75)'); halo.addColorStop(0.25, 'rgba(255,200,90,0.28)'); halo.addColorStop(1, 'rgba(255,170,60,0)');
        c.fillStyle = halo; c.beginPath(); c.arc(cx, SY, 110, 0, 7); c.fill();
        // long, slowly turning light rays
        c.translate(cx, SY); c.rotate(T * 0.05);
        for (let r = 0; r < 16; r++) {
          const len = r % 2 ? 70 : 105, a = r * Math.PI / 8, rg = c.createLinearGradient(0, 0, Math.cos(a) * len, Math.sin(a) * len);
          rg.addColorStop(0, 'rgba(255,240,190,0.55)'); rg.addColorStop(1, 'rgba(255,220,130,0)');
          c.strokeStyle = rg; c.lineWidth = r % 2 ? 2 : 3.5; c.lineCap = 'round';
          c.beginPath(); c.moveTo(Math.cos(a) * 20, Math.sin(a) * 20); c.lineTo(Math.cos(a) * len, Math.sin(a) * len); c.stroke();
        }
        c.restore();
        // glowing white-hot core
        c.save(); c.globalCompositeOperation = 'source-over';
        const core = c.createRadialGradient(cx, SY, 0, cx, SY, 22);
        core.addColorStop(0, '#ffffff'); core.addColorStop(0.45, '#fffbe0'); core.addColorStop(0.75, 'rgba(255,236,150,0.9)'); core.addColorStop(1, 'rgba(255,220,120,0)');
        c.fillStyle = core; c.beginPath(); c.arc(cx, SY, 22, 0, 7); c.fill();
        c.restore();
      };
      g.save(); g.beginPath(); g.rect(0, 0, SW, 262); g.clip();
      g.globalAlpha = 0.6; g.drawImage(cloud as HTMLCanvasElement, 0, 0, SW * 2, 524, 0, 0, SW, 262);
      g.globalAlpha = 1; disc(g, SX);
      g.restore();
      if (gl) disc(gl, SX + SW);
      if (gr) disc(gr, SX - SW);
      // drifting cloud puffs
      for (let i = 0; i < 6; i++) {
        const cx = ((i * 380 + T * (6 + i * 2)) % 2100) - 200, cy = 40 + i * 32 + Math.sin(i * 2.1) * 20;
        for (let k = 0; k < 4; k++) {
          const px = cx + k * 46, r = 60 + k * 8, rg = g.createRadialGradient(px, cy, 0, px, cy, r);
          rg.addColorStop(0, 'rgba(255,240,220,0.3)'); rg.addColorStop(1, 'rgba(255,240,220,0)');
          g.fillStyle = rg; g.beginPath(); g.ellipse(px, cy, r * 1.5, r * 0.5, 0, 0, 7); g.fill();
        }
      }
      // ship drawn last, untouched, in front of sun & clouds
      g.save(); g.beginPath(); poly.forEach((p, i) => (i ? g.lineTo(p[0], p[1]) : g.moveTo(p[0], p[1]))); g.closePath(); g.clip(); d(0, 0, SW, SH, 0, 0, SW, SH); g.restore();
      flecks.forEach(f => {
        const x = (((f.x - T * 14 * f.v) % SW) + SW) % SW, y = f.y + Math.sin(T * 1.4 + f.p) * 3, a = Math.max(0, Math.sin(T * 1.8 * f.v + f.p));
        if (inShip(x, y)) return;
        g.fillStyle = `rgba(255,255,255,${a * 0.5})`; g.beginPath(); g.ellipse(x, y, f.w, 1 + f.w * 0.06, 0, 0, 7); g.fill();
      });
    };
    tick();
    if (!still) timer = window.setInterval(tick, 40); else { const t2 = window.setTimeout(tick, 400); return () => window.clearTimeout(t2); }
    return () => window.clearInterval(timer);
  }, [ref, night]);
}

function AuthScreen({ host, mode }: { host: VesselEmail; mode: AuthPageMode }): React.ReactElement {
  injectFuturisticTheme();
  const cvRef = React.useRef<HTMLCanvasElement>(null);
  const wheelRef = React.useRef<SVGGElement>(null);
  const lRef = React.useRef<HTMLCanvasElement>(null);
  const rRef = React.useRef<HTMLCanvasElement>(null);
  const mRef = React.useRef(0);
  const [voyage, setVoyage] = React.useState(false);
  const [tab, setTab] = React.useState<'signin' | 'signup'>('signin');
  const [vp, setVp] = React.useState<[number, number]>(typeof window !== 'undefined' ? [window.innerWidth, window.innerHeight] : [1440, 900]);
  React.useEffect(() => {
    const on = (): void => setVp([window.innerWidth, window.innerHeight]);
    window.addEventListener('resize', on);
    return () => window.removeEventListener('resize', on);
  }, []);
  const isNight = host.state.themeMode === 'night';
  React.useEffect(() => {
    if (reduceMotion()) return undefined;
    const id = window.setInterval(() => {
      const t = Date.now() / 1000;
      if (wheelRef.current) wheelRef.current.style.transform = `rotate(${Math.sin(t * 0.6) * 25 + t * 4}deg)`;
    }, 40);
    return () => window.clearInterval(id);
  }, []);

  const isLogout = mode === 'logout';
  const signUp = tab === 'signup' && !isLogout;
  const displayName = host.props.userDisplayName || '';
  const tenant = tenantHost(host) || 'contoso.sharepoint.com';
  const [w, h] = vp;
  const wide = w >= 1000;
  // Wide layout: scale to the full height so the card and feature row are never cropped; mirrored side canvases fill the width.
  const k = wide ? h / SH : Math.max(w / SW, h / SH);
  const tx = SW * k <= w ? (w - SW * k) / 2 : Math.max(w - SW * k, Math.min((w - SW * k) / 2, w - 1650 * k));
  const ty = wide ? 0 : (h - SH * k) / 2;
  mRef.current = Math.max(0, tx / k, (w - tx - SW * k) / k);
  useSceneAnimation(cvRef, isNight, lRef, rRef, mRef);
  const blurb = isLogout ? 'Your Vessel DMS session has been closed. Return to SharePoint to start a new session.'
    : signUp ? "Accounts use your organisation's Microsoft 365 identity. We'll send you to Microsoft to finish."
    : `Continue with your SharePoint account${displayName ? `, ${displayName}` : ''} to access vessel documents and workflows.`;
  const cta = isLogout ? 'Return to sign in' : signUp ? 'Sign up with SharePoint' : 'Continue with SharePoint';
  const on = 'linear-gradient(180deg,#3a8cf0,#0a44b0)';
  const tabBtn = (t: 'signin' | 'signup'): React.CSSProperties => ({ border: 'none', cursor: 'pointer', fontFamily: vdmsFont.ui, fontSize: 20, fontWeight: 700, padding: '13px 0', borderRadius: 13, background: tab === t ? on : 'transparent', color: tab === t ? '#fff' : '#173a70', boxShadow: tab === t ? '0 6px 14px rgba(17,80,190,0.35)' : 'none' });

  const card = (
    <div className="vdms-pop" style={{ boxSizing: 'border-box', padding: '0 60px', borderRadius: 52, background: 'linear-gradient(180deg,#ffffff 0%,#eaf4ff 60%,#cfe6ff 100%)', border: '1px solid #fff', boxShadow: '0 30px 70px rgba(2,20,52,0.35), inset 0 1px 0 #fff', textAlign: 'center', overflow: 'hidden', display: 'flex', flexDirection: 'column', justifyContent: 'center', color: '#0b2247', fontFamily: vdmsFont.ui, ...(wide ? { position: 'absolute', left: 1002, top: 84, width: 608, height: 751 } : { width: 'min(440px, 92vw)', padding: '28px 26px', zoom: 0.78 } as React.CSSProperties) }}>
      <svg width="104" height="104" viewBox="0 0 76 76" style={{ display: 'block', margin: '0 auto' }} fill="none" stroke="#1652b8" strokeWidth="3.6" strokeLinecap="round">
        <g ref={wheelRef} style={{ transformOrigin: '38px 34px' }}>
          <circle cx="38" cy="34" r="22" /><circle cx="38" cy="34" r="7" />
          <path d="M38 12V4M38 56V64M16 34H8M60 34H68M22.4 18.4L16.7 12.7M53.6 49.6L59.3 55.3M22.4 49.6L16.7 55.3M53.6 18.4L59.3 12.7" />
        </g>
        <path d="M14 68c8-6 16-6 24 0s16 6 24 0" stroke="#2a8fe0" />
      </svg>
      <h1 style={{ margin: 0, fontFamily: vdmsFont.display, fontSize: 56, fontWeight: 800, letterSpacing: '-0.03em', lineHeight: 1.1 }}><span style={{ color: '#0b2247' }}>Vessel</span> <span style={{ color: '#1652b8' }}>DMS</span></h1>
      <div style={{ marginTop: 8, fontSize: 21, fontWeight: 500, color: '#23406e' }}>Your Vessel Documents, Always Accessible</div>
      {!isLogout && (
        <div role="tablist" style={{ marginTop: 34, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 4, padding: 6, borderRadius: 18, background: '#fff', border: '1px solid rgba(190,212,240,0.9)' }}>
          <button type="button" role="tab" aria-selected={tab === 'signin'} onClick={() => setTab('signin')} style={tabBtn('signin')}>Sign in</button>
          <button type="button" role="tab" aria-selected={tab === 'signup'} onClick={() => setTab('signup')} style={tabBtn('signup')}>Sign up</button>
        </div>
      )}
      <p style={{ margin: '20px auto 22px', maxWidth: 440, minHeight: 56, fontSize: 19, lineHeight: 1.5, color: '#2b4874', fontWeight: 500 }}>{blurb}</p>
      <button type="button" onClick={() => (isLogout ? goToSignIn(host) : setVoyage(true))} style={{ width: '100%', minHeight: 70, border: 'none', borderRadius: 16, cursor: 'pointer', fontFamily: vdmsFont.ui, whiteSpace: 'nowrap', fontSize: 21, fontWeight: 700, color: '#fff', background: on, boxShadow: '0 12px 28px rgba(10,80,220,0.5), inset 0 1px 0 rgba(255,255,255,0.35)', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 14 }}>
        <span style={{ width: 36, height: 36, borderRadius: 9, background: '#fff', position: 'relative', display: 'inline-block', overflow: 'hidden' }}>
          <span style={{ position: 'absolute', left: 11, top: 2, width: 22, height: 22, borderRadius: '50%', background: '#1a9ba1' }} />
          <span style={{ position: 'absolute', left: 17, top: 13, width: 19, height: 19, borderRadius: '50%', background: '#37c6d0' }} />
          <span style={{ position: 'absolute', left: 3, top: 10, width: 19, height: 19, borderRadius: 6, background: '#0e6b73', color: '#fff', fontSize: 13, fontWeight: 800, lineHeight: '19px' }}>S</span>
        </span>
        <span>{cta}</span><span style={{ fontSize: 26, lineHeight: 1 }}>›</span>
      </button>
      <div title="Signing in to this SharePoint tenant" style={{ position: 'relative', width: 390, maxWidth: '100%', margin: '22px auto 0' }}>
        <span style={{ position: 'absolute', left: 18, top: '50%', transform: 'translateY(-50%)', width: 16, height: 21, border: '2px solid #23406e', borderRadius: 2, boxSizing: 'border-box' }} />
        <div style={{ boxSizing: 'border-box', padding: '16px 46px 16px 52px', borderRadius: 16, border: '1px solid rgba(190,212,240,0.95)', background: '#fff', fontSize: 18, fontWeight: 500, color: '#0b2247', textAlign: 'left', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{tenant}</div>
      </div>
      <div style={{ margin: '18px auto 0', display: 'inline-flex', whiteSpace: 'nowrap', alignItems: 'center', gap: 12, padding: '10px 24px', borderRadius: 999, background: '#fff', border: '1px solid rgba(190,212,240,0.9)', fontFamily: vdmsFont.mono, fontSize: 16, color: '#1c3a66' }}><span className="vdms-pulse" style={{ width: 9, height: 9, borderRadius: '50%', background: '#12a05c' }} />All services operational</div>
      <div style={{ marginTop: 22, fontSize: 17, color: '#2b4874', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10 }}><span style={{ width: 12, height: 9, border: '2px solid #23406e', borderRadius: 2, position: 'relative', top: 2 }} />Secure access is managed by your organization.</div>
    </div>
  );

  return (
    <div className="vessel-dms-auth" data-vessel-theme={host.state.themeMode} style={{ position: 'fixed', inset: 0, zIndex: 99999, overflow: 'hidden', background: '#062a5c', fontFamily: vdmsFont.ui, display: wide ? 'block' : 'flex', alignItems: 'center', justifyContent: 'center' }}>
      {voyage && <SignInVoyage onDone={() => goToSignIn(host)} />}
      <div aria-hidden={!wide} style={{ position: 'absolute', left: 0, top: 0, width: SW, height: SH, transformOrigin: '0 0', transform: `translate(${tx}px,${ty}px) scale(${k})`, background: '#062a5c' }}>
        {wide && <canvas ref={lRef} width={SW * 2} height={SH * 2} style={{ position: 'absolute', left: -SW, top: 0, width: SW, height: SH, filter: isNight ? 'brightness(0.55) saturate(1.1)' : 'none' }} />}
        {wide && <canvas ref={rRef} width={SW * 2} height={SH * 2} style={{ position: 'absolute', left: SW, top: 0, width: SW, height: SH, filter: isNight ? 'brightness(0.55) saturate(1.1)' : 'none' }} />}
        <canvas ref={cvRef} width={SW * 2} height={SH * 2} style={{ position: 'absolute', left: 0, top: 0, width: SW, height: SH, filter: isNight ? 'brightness(0.55) saturate(1.1)' : 'none' }} />
        {wide && (
          card
        )}
      </div>
      {!wide && <div style={{ position: 'relative', zIndex: 1, maxHeight: '100%', overflowY: 'auto', padding: '12px 0' }}>{card}</div>}
    </div>
  );
}
