// Maritime futuristic skin — DESIGN ONLY.
// A single extra <style> tag layered on top of AppLayout.tsx's
// injectFullScreenStyles(). It only re-maps CSS custom properties and adds
// presentation rules (glass, radii, fonts, background, motion). No handler,
// state, API call, permission or DOM order is touched.
//
// Settings → Color Management keeps working: every --clay-* value that
// applyColorTheme() writes (bg / text / accent / hover and their derived
// surfaces, gradients, glow) is deliberately NOT overridden here.

import { skySeaBackgroundImage } from './skySeaBackground';
import { appBackgroundImage } from './appBackground';

const STYLE_ID = 'vessel-dms-futuristic';
const FONT_ID = 'vessel-dms-futuristic-fonts';

export const vdmsFont = {
  display: "'Sora', 'Segoe UI Variable', 'Segoe UI', sans-serif",
  ui: "'Manrope', 'Segoe UI Variable', 'Segoe UI', sans-serif",
  mono: "'JetBrains Mono', Consolas, monospace",
};

export const vdmsMotion = {
  spring: 'cubic-bezier(.2,.9,.25,1.15)',
  glide: 'cubic-bezier(.2,.8,.2,1)',
};

export function injectFuturisticTheme(): void {
  if (typeof document === 'undefined') return;

  if (!document.getElementById(FONT_ID)) {
    const link = document.createElement('link');
    link.id = FONT_ID;
    link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Manrope:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap';
    document.head.appendChild(link);
  }

  let style = document.getElementById(STYLE_ID) as HTMLStyleElement | null;
  if (style) return; // static content — inject once
  style = document.createElement('style');
  style.id = STYLE_ID;
  document.head.appendChild(style);

  style.textContent = `
    /* ── Neutral + glass tokens (not admin-controlled) ─────────────────── */
    html [data-vessel-theme="light"] {
      --vdms-surface: rgba(255,255,255,0.86);
      --vdms-surface-alt: rgba(236,247,255,0.9);
      --vdms-border: rgba(16,84,138,0.16);
      --vdms-border-soft: rgba(16,84,138,0.09);
      --vdms-text: #08243a;
      --vdms-text-secondary: #34536a;
      --vdms-text-muted: #4a6a82;
      --vdms-text-faint: #6b879c;
      --vdms-toggle-active-bg: #0a7ea8;
      --vdms-toggle-active-text: #ffffff;
      --vdms-glass: rgba(255,255,255,0.62);
      --vdms-glass-strong: rgba(255,255,255,0.92);
      --vdms-line: rgba(16,84,138,0.14);
      --vdms-line-strong: rgba(16,84,138,0.26);
      --vdms-field: rgba(255,255,255,0.72);
      --vdms-shadow: 0 24px 60px rgba(20,80,130,0.18);
      --vdms-focus: rgba(10,126,168,0.7);
      --vdms-focus-soft: rgba(10,126,168,0.16);
      --clay-shadow-raised: 0 12px 32px rgba(20,80,130,0.14), inset 0 1px 0 rgba(255,255,255,0.7);
      --clay-shadow-raised-hover: 0 18px 42px rgba(20,80,130,0.2), inset 0 1px 0 rgba(255,255,255,0.8);
      --clay-shadow-icon: inset 0 1px 0 rgba(255,255,255,0.75), inset 0 0 0 1px rgba(10,126,168,0.12);
      --clay-pill-active-bg: #d3f4e6; --clay-pill-active-text: #0b6b49; --clay-pill-active-shadow: none;
      --clay-pill-warn-bg: #fdf0d2;   --clay-pill-warn-text: #8a5a00;   --clay-pill-warn-shadow: none;
      --clay-pill-danger-bg: #fde0e4; --clay-pill-danger-text: #b3243c; --clay-pill-danger-shadow: none;
    }
    html [data-vessel-theme="night"] {
      --vdms-surface: rgba(9,34,56,0.82);
      --vdms-surface-alt: rgba(12,44,70,0.85);
      --vdms-border: rgba(140,210,240,0.18);
      --vdms-border-soft: rgba(140,210,240,0.1);
      --vdms-text: #eaf6fd;
      --vdms-text-secondary: #b2cadb;
      --vdms-text-muted: #9ab4c6;
      --vdms-text-faint: #7f9bb0;
      --vdms-toggle-active-bg: #34d5ea;
      --vdms-toggle-active-text: #02202b;
      --vdms-glass: rgba(9,34,56,0.6);
      --vdms-glass-strong: rgba(7,28,47,0.92);
      --vdms-line: rgba(140,210,240,0.16);
      --vdms-line-strong: rgba(140,210,240,0.28);
      --vdms-field: rgba(3,20,35,0.5);
      --vdms-shadow: 0 24px 60px rgba(0,8,20,0.55);
      --vdms-focus: rgba(52,213,234,0.75);
      --vdms-focus-soft: rgba(52,213,234,0.18);
      --clay-shadow-raised: 0 14px 36px rgba(0,8,20,0.5), inset 0 1px 0 rgba(255,255,255,0.05);
      --clay-shadow-raised-hover: 0 20px 44px rgba(0,8,20,0.6), inset 0 1px 0 rgba(255,255,255,0.07);
      --clay-shadow-icon: inset 0 1px 0 rgba(255,255,255,0.08), inset 0 0 0 1px rgba(140,210,240,0.14);
      --clay-pill-active-bg: rgba(67,224,164,0.16); --clay-pill-active-text: #6ff0bd; --clay-pill-active-shadow: none;
      --clay-pill-warn-bg: rgba(255,200,103,0.16);  --clay-pill-warn-text: #ffd27f;  --clay-pill-warn-shadow: none;
      --clay-pill-danger-bg: rgba(255,122,138,0.16); --clay-pill-danger-text: #ff98a5; --clay-pill-danger-shadow: none;
    }

    /* ── Type ─────────────────────────────────────────────────────────── */
    .vessel-dms-app, .vessel-dms-auth { font-family: ${vdmsFont.ui} !important; -webkit-font-smoothing: antialiased; }
    .vessel-dms-app h1, .vessel-dms-app h2, .vessel-dms-app h3, .vessel-dms-auth h1 { font-family: ${vdmsFont.display}; letter-spacing: -0.02em; }

    /* ── Sea & sky backdrop ───────────────────────────────────────────── */
    /* In-app backdrop = the sign-in scene photo (ship & sea), cropped clean of the login card. */
    .vessel-dms-app[data-vessel-theme="light"] {
      background:
        linear-gradient(rgba(255,255,255,0.38), rgba(255,255,255,0.38)),
        url("${appBackgroundImage}") center 50% / cover no-repeat,
        #2f86d6 !important;
    }
    .vessel-dms-app[data-vessel-theme="night"] {
      background:
        linear-gradient(180deg, rgba(3,20,35,0.62) 0%, rgba(3,20,35,0.78) 100%),
        url("${appBackgroundImage}") center 50% / cover no-repeat,
        #031627 !important;
    }
    .vessel-dms-auth[data-vessel-theme="light"] {
      background:
        radial-gradient(70% 50% at 18% 0%, rgba(255,255,255,0.35), transparent 70%),
        linear-gradient(180deg, rgba(6,70,150,0.18) 0%, rgba(255,255,255,0) 38%, rgba(2,50,110,0.22) 100%),
        url("${skySeaBackgroundImage}") center 55% / cover no-repeat,
        #2f86d6 !important;
    }
    .vessel-dms-auth[data-vessel-theme="night"] {
      background:
        radial-gradient(120% 70% at 30% -10%, rgba(30,140,200,0.5) 0%, rgba(30,140,200,0) 60%),
        radial-gradient(80% 60% at 100% 100%, rgba(10,90,150,0.45) 0%, rgba(10,90,150,0) 70%),
        linear-gradient(180deg, rgba(3,20,35,0.62) 0%, rgba(3,20,35,0.78) 100%),
        url("${skySeaBackgroundImage}") center 60% / cover no-repeat,
        #031627 !important;
    }
    .vessel-dms-auth[data-vessel-theme="night"] .vdms-auth-photo { background-image: url("${skySeaBackgroundImage}"); filter: brightness(0.42) saturate(1.2); }
    .vessel-dms-auth[data-vessel-theme="light"] .vdms-auth-photo { background-image: url("${skySeaBackgroundImage}"); }
    /* Ambient layer: drifting light (sky) / caustic rays (sea). Sits under all content. */
    .vessel-dms-app::before, .vessel-dms-auth::before {
      content: ''; position: absolute; inset: -10% -20%; z-index: -1; pointer-events: none;
      animation: vdms-rays 26s linear infinite;
    }
    .vessel-dms-app[data-vessel-theme="light"]::before, .vessel-dms-auth[data-vessel-theme="light"]::before {
      background: radial-gradient(40% 18% at 20% 20%, rgba(255,255,255,0.55), transparent 70%),
                  radial-gradient(35% 14% at 70% 12%, rgba(255,255,255,0.45), transparent 70%);
      animation: vdms-cloud 90s ease-in-out infinite alternate;
    }
    .vessel-dms-app[data-vessel-theme="night"]::before, .vessel-dms-auth[data-vessel-theme="night"]::before {
      background: repeating-linear-gradient(105deg, rgba(120,220,255,0.05) 0 3px, transparent 3px 113px);
      -webkit-mask-image: linear-gradient(to bottom, #000, transparent 85%); mask-image: linear-gradient(to bottom, #000, transparent 85%);
    }
    .vessel-dms-app { isolation: isolate; }

    /* ── Bright 3D sidebar ─────────────────────────────────────────────── */
    html [data-vessel-theme="light"] .vessel-dms-sidebar {
      background: linear-gradient(170deg, #ffffff 0%, #f1f9ff 40%, #dff0fd 100%) !important;
      border: 1px solid rgba(255,255,255,0.95);
      box-shadow: 0 30px 70px rgba(20,80,130,0.28), 0 8px 20px rgba(20,80,130,0.14), inset 0 1px 0 #fff, inset 0 -6px 0 rgba(10,90,150,0.06) !important;
    }
    html [data-vessel-theme="night"] .vessel-dms-sidebar {
      background: linear-gradient(170deg, #0f3a5f 0%, #0a2a48 45%, #061c33 100%) !important;
      border: 1px solid rgba(140,210,240,0.22);
      box-shadow: 0 30px 70px rgba(0,8,20,0.6), 0 8px 20px rgba(0,8,20,0.4), inset 0 1px 0 rgba(255,255,255,0.12) !important;
    }
    .vessel-dms-sidebar-nav button { transform-origin: left center; transition: transform .35s ${vdmsMotion.glide}, box-shadow .3s ease, background .3s ease !important; }
    html .vessel-dms-app .vessel-dms-sidebar-nav button:hover:not(:disabled) {
      /* Design Guidelines §5: hover = 3px nudge right + soft fill. No tilt / Z-lift, so rows never drift up or down. */
      transform: translateX(3px) !important;
      outline: none !important;
    }
    /* 3D bevel on every nav icon tile (keeps each module's own colour) */
    .vessel-dms-sidebar-nav button > span:first-child {
      box-shadow: 0 6px 14px rgba(10,60,110,0.28), inset 0 2px 0 rgba(255,255,255,0.45), inset 0 -3px 0 rgba(0,0,0,0.2) !important;
      border-radius: 14px !important;
    }
    .vessel-dms-sidebar {
      border-radius: 28px;
      margin: 14px 0 14px 14px; height: calc(100vh - 28px) !important;
      transition: width .35s ${vdmsMotion.glide}, min-width .35s ${vdmsMotion.glide}, max-width .35s ${vdmsMotion.glide}, transform .35s ${vdmsMotion.glide}, opacity .25s ease !important;
    }
    .vessel-dms-sidebar > div { background: transparent !important; border-color: var(--vdms-line) !important; }
    .vessel-dms-sidebar button { border-radius: 16px; font-size: 15px; }
    @media (max-width: 1024px) { .vessel-dms-sidebar { top: 14px !important; left: 14px !important; margin: 0; } }

    /* ── Floating top bar ─────────────────────────────────────────────── */
    /* Sticky glass top bar (colours come from inline var(--vdms-*) in AppLayout) */
    .vessel-dms-topbar {
      margin: 14px 14px 0; height: 72px !important; border-radius: 22px;
      border: 1px solid var(--vdms-line) !important;
      backdrop-filter: blur(22px) saturate(1.4); -webkit-backdrop-filter: blur(22px) saturate(1.4);
      position: relative; z-index: 5;
    }
    .vessel-dms-topbar, .vessel-dms-topbar button { font-family: ${vdmsFont.ui}; }

    /* ── Dashboard hero + glass cards ─────────────────────────────────── */
    html [data-vessel-theme="light"] .vdms-hero-ovl { background: linear-gradient(90deg, rgba(236,247,255,0.95) 0%, rgba(236,247,255,0.7) 55%, rgba(236,247,255,0.05) 100%); }
    html [data-vessel-theme="night"] .vdms-hero-ovl { background: linear-gradient(90deg, rgba(3,20,35,0.92) 0%, rgba(3,20,35,0.6) 55%, rgba(3,20,35,0.15) 100%); }
    html [data-vessel-theme="night"] .vdms-hero { background-color: #031423; filter: none; }
    .vessel-dms-app div[style*="var(--vdms-surface)"][style*="border-radius"],
    .vessel-dms-app div[style*="--clay-surface"][style*="--clay-shadow-raised,"] {
      backdrop-filter: blur(16px) saturate(1.25); -webkit-backdrop-filter: blur(16px) saturate(1.25);
    }
    .vessel-dms-app table th { font-family: ${vdmsFont.mono}; letter-spacing: 0.08em; }
    /* ── Typography scale (design: body 16–17, table cells 15, labels 14, meta 12) ──
       Pages set inline px sizes; these remap them without touching any JSX. */
    .vessel-dms-app [style*="font-size: 9px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 11px !important; }
    .vessel-dms-app [style*="font-size: 10px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 11px !important; }
    .vessel-dms-app [style*="font-size: 11px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 12px !important; }
    .vessel-dms-app [style*="font-size: 11.5px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 13px !important; }
    .vessel-dms-app [style*="font-size: 12px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 14px !important; }
    .vessel-dms-app [style*="font-size: 12.5px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 14px !important; }
    .vessel-dms-app [style*="font-size: 13px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 15px !important; }
    .vessel-dms-app [style*="font-size: 14px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 16px !important; }
    .vessel-dms-app [style*="font-size: 15px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 16px !important; }
    .vessel-dms-app [style*="font-size: 16px"]:not(i):not([data-icon-name]):not([data-no-scale]) { font-size: 17px !important; }
    .vessel-dms-app h1 { font-size: 28px !important; }
    .vessel-dms-app h2 { font-size: 22px !important; }
    .vessel-dms-app h3 { font-size: 19px !important; }
    .vessel-dms-app [style*="font-family"]:not(i):not([data-icon-name]):not([data-no-scale]):not([style*="Sora"]):not([style*="JetBrains"]) { font-family: ${vdmsFont.ui} !important; }
    .vessel-dms-app table td { font-size: 15px; }
    .vessel-dms-app table th { font-size: 12px !important; }

    /* ── Text colours: hard-coded slate greys → design ink (light mode only;
       night keeps the var(--vdms-*) values, which already flip) ───────── */
    html [data-vessel-theme="light"] [style*="color: rgb(15, 23, 42)"]:not([style*="background: rgb"]) { color: #08243a !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(30, 41, 59)"]:not([style*="background: rgb"]) { color: #08243a !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(51, 65, 85)"]:not([style*="background: rgb"]) { color: #34536a !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(71, 85, 105)"]:not([style*="background: rgb"]) { color: #34536a !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(100, 116, 139)"]:not([style*="background: rgb"]) { color: #4a6a82 !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(148, 163, 184)"]:not([style*="background: rgb"]) { color: #6b879c !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(2, 132, 199)"]:not([style*="background: rgb"]) { color: #0a7ea8 !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(3, 105, 161)"]:not([style*="background: rgb"]) { color: #0b5f8a !important; }
    html [data-vessel-theme="light"] [style*="color: rgb(7, 89, 133)"]:not([style*="background: rgb"]) { color: #0a4a73 !important; }


    /* ── Inputs ───────────────────────────────────────────────────────── */
    .vessel-dms-app input, .vessel-dms-app select, .vessel-dms-app textarea { border-radius: 12px; font-family: ${vdmsFont.ui}; }
    html [data-vessel-theme="night"] input, html [data-vessel-theme="night"] select, html [data-vessel-theme="night"] textarea {
      background: rgba(3,20,35,0.6) !important; color: #eaf6fd !important; border-color: rgba(140,210,240,0.28) !important;
    }
    html [data-vessel-theme="night"] input::placeholder, html [data-vessel-theme="night"] textarea::placeholder { color: #7f9bb0 !important; }
    .vessel-dms-app input:focus, .vessel-dms-app select:focus, .vessel-dms-app textarea:focus {
      outline: none; border-color: var(--clay-accent) !important; box-shadow: 0 0 0 4px var(--vdms-focus-soft) !important;
    }

    /* ── Interaction feedback (replaces the orange outline) ───────────── */
    html .vessel-dms-app button, html .vessel-dms-app a, html .vessel-dms-app [role="button"] {
      transition: filter .2s ease, transform .22s ${vdmsMotion.glide}, box-shadow .22s ease, background .25s ease;
    }
    html .vessel-dms-app button:hover:not(:disabled), html .vessel-dms-app a:hover, html .vessel-dms-app [role="button"]:hover {
      filter: brightness(1.04) saturate(1.06) !important;
      outline: 2px solid var(--vdms-focus) !important; outline-offset: 2px;
      box-shadow: 0 0 0 4px var(--vdms-focus-soft) !important;
    }
    html .vessel-dms-app button:focus-visible, html .vessel-dms-app a:focus-visible, html .vessel-dms-app [role="button"]:focus-visible {
      outline: 3px solid var(--vdms-focus) !important; outline-offset: 2px;
    }

    /* ── Popups: blurred overlay + spring entrance ────────────────────── */
    .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] {
      backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px);
      animation: vdms-fade .25s ease both;
    }
    .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] > div:only-child {
      border-radius: 24px; animation: vdms-pop .45s ${vdmsMotion.spring} both;
    }

    /* ── Motion primitives (also used by AuthPage.tsx) ────────────────── */
    @keyframes vdms-bob { 0%, 100% { transform: translateY(0) } 50% { transform: translateY(-14px) } }
    @keyframes vdms-cube { from { transform: rotateX(-18deg) rotateY(0deg) } to { transform: rotateX(-18deg) rotateY(360deg) } }
    @keyframes vdms-sheen3d { 0% { transform: translateX(-130%) skewX(-18deg) } 60%, 100% { transform: translateX(260%) skewX(-18deg) } }
    .vdms-bob { animation: vdms-bob 7s ease-in-out infinite; }
    .vdms-cube { position: relative; transform-style: preserve-3d; animation: vdms-cube 12s linear infinite; }
    .vdms-sheen3d { position: absolute; top: 0; bottom: 0; left: 0; width: 45%; background: linear-gradient(100deg, transparent, rgba(255,255,255,0.28), transparent); animation: vdms-sheen3d 6s ease-in-out infinite; pointer-events: none; }
    @keyframes vdms-helm-sway { 0%, 100% { transform: rotate(-30deg) } 50% { transform: rotate(30deg) } }
    @keyframes vdms-helm-spin { from { transform: rotate(0) } to { transform: rotate(-540deg) } }
    @keyframes vdms-cube-out { from { transform: translateX(0) rotateY(0) } to { transform: translateX(-100%) rotateY(90deg) } }
    @keyframes vdms-cube-in { from { transform: translateX(100%) rotateY(-90deg) } to { transform: translateX(0) rotateY(0) } }
    @keyframes vdms-sail { 0%, 100% { transform: scale(1.08) translateX(-2%) rotate(-.5deg) } 50% { transform: scale(1.08) translateX(2%) translateY(-8px) rotate(.5deg) } }
    @keyframes vdms-cap { from { opacity: 0; transform: translate(-50%, 14px) } to { opacity: 1; transform: translate(-50%, 0) } }
    @keyframes vdms-walk { 0%, 45% { transform: scale(1); opacity: 1 } 100% { transform: scale(2.8); opacity: 0 } }
    @keyframes vdms-leaf-l { 0%, 15% { transform: rotateY(0) } 60%, 100% { transform: rotateY(-108deg) } }
    @keyframes vdms-leaf-r { 0%, 15% { transform: rotateY(0) } 60%, 100% { transform: rotateY(108deg) } }
    @keyframes vdms-deck-l { 0%, 25% { transform: rotateY(0) } 100% { transform: rotateY(-100deg) } }
    @keyframes vdms-deck-r { 0%, 25% { transform: rotateY(0) } 100% { transform: rotateY(100deg) } }
    @keyframes vdms-plate { 0%, 30% { opacity: 1; transform: translate(-50%,-50%) scale(1) } 70%, 100% { opacity: 0; transform: translate(-50%,-50%) scale(1.06) } }
    @keyframes vdms-room-in { from { opacity: 0; transform: perspective(1200px) rotateY(-10deg) translateZ(-120px) } to { opacity: 1; transform: none } }
    @keyframes vdms-voyage { from { transform: translateX(-45vw) } to { transform: translateX(110vw) } }
    @keyframes vdms-ship-in { from { opacity: 0; transform: translateX(-30%) scale(.9) } to { opacity: 1; transform: none } }
    @keyframes vdms-pull { 0%, 70%, 100% { transform: translateY(0) } 78%, 92% { transform: translateY(-16px) } }
    @keyframes vdms-fade { from { opacity: 0 } to { opacity: 1 } }
    @keyframes vdms-cube-spin { 0% { transform: rotateX(-18deg) rotateY(0deg) } 100% { transform: rotateX(-18deg) rotateY(360deg) } }
    @keyframes vdms-pop { from { opacity: 0; transform: translateY(20px) scale(.95) } to { opacity: 1; transform: none } }
    @keyframes vdms-sonar { 0% { transform: scale(1); opacity: .75 } 100% { transform: scale(3.2); opacity: 0 } }
    @keyframes vdms-shimmer { 0% { transform: translateX(-120%) } 55%, 100% { transform: translateX(320%) } }
    @keyframes vdms-kenburns { from { transform: scale(1.03) } to { transform: scale(1.12) translate(-1.5%, -1%) } }
    @keyframes vdms-rays { from { transform: translateX(0) } to { transform: translateX(113px) } }
    @keyframes vdms-cloud { from { transform: translateX(-6%) } to { transform: translateX(6%) } }
    @keyframes vdms-pulse { 0%, 100% { box-shadow: 0 0 0 0 rgba(67,224,164,0.5) } 50% { box-shadow: 0 0 0 7px rgba(67,224,164,0) } }
    .vdms-sonar { position: absolute; inset: 0; border-radius: 24px; border: 1.5px solid var(--clay-accent, #0a7ea8); animation: vdms-sonar 3.3s cubic-bezier(.1,.6,.3,1) infinite; pointer-events: none; }
    .vdms-shimmer { position: absolute; top: 0; bottom: 0; left: 0; width: 40%; background: linear-gradient(100deg, transparent, rgba(255,255,255,0.38), transparent); animation: vdms-shimmer 4s ease-in-out infinite; pointer-events: none; }
    .vdms-pop { animation: vdms-pop .8s ${vdmsMotion.spring} both; }
    .vdms-pulse { animation: vdms-pulse 2.4s ease-in-out infinite; }
    .vdms-auth-photo { position: absolute; inset: 0; background-size: cover; background-position: center 40%; animation: vdms-kenburns 40s ease-in-out infinite alternate; }

    /* ── Responsive: TV / desktop / laptop / tablet / phone ───────────────
       The page-level layouts switch on JS breakpoints (responsive.ts: phone ≤767,
       tablet ≤1024); everything below is the global safety net + large-screen scaling. */
    .vessel-dms-app { --vdms-z: 1; }
    @media (min-width: 1920px) { .vessel-dms-app { --vdms-z: 1.15; } }
    @media (min-width: 2400px) { .vessel-dms-app { --vdms-z: 1.4; } }
    @media (min-width: 3000px) { .vessel-dms-app { --vdms-z: 1.8; } }
    @media (min-width: 3800px) { .vessel-dms-app { --vdms-z: 2.2; } }
    html .vessel-dms-app {
      zoom: var(--vdms-z);
      width: calc(100vw / var(--vdms-z)) !important;
      height: calc(100vh / var(--vdms-z)) !important; min-height: 0 !important;
    }
    @supports (height: 100dvh) { html .vessel-dms-app { height: calc(100dvh / var(--vdms-z)) !important; } }
    html .vessel-dms-app .vessel-dms-sidebar { height: calc(100vh / var(--vdms-z) - 28px) !important; }
    @supports (height: 100dvh) { html .vessel-dms-app .vessel-dms-sidebar { height: calc(100dvh / var(--vdms-z) - 28px) !important; } }
    .vessel-dms-app img, .vessel-dms-app video, .vessel-dms-app canvas, .vessel-dms-app iframe { max-width: 100%; }
    .vessel-dms-app input:not([type="checkbox"]):not([type="radio"]), .vessel-dms-app select, .vessel-dms-app textarea { max-width: 100%; box-sizing: border-box; }
    /* Dialogs never exceed the screen; their content scrolls instead */
    .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] > div:only-child {
      max-width: calc(100vw / var(--vdms-z) - 16px); max-height: calc(100vh / var(--vdms-z) - 16px); overflow: auto; box-sizing: border-box;
    }
    @supports (height: 100dvh) { .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] > div:only-child { max-height: calc(100dvh / var(--vdms-z) - 16px); } }
    /* Touch screens: comfortable tap targets */
    @media (pointer: coarse) {
      .vessel-dms-app button, .vessel-dms-app select, .vessel-dms-app input:not([type="checkbox"]):not([type="radio"]) { min-height: 40px; }
    }
    @media (max-width: 1024px) {
      .vessel-dms-app table { -webkit-overflow-scrolling: touch; }
      .vessel-dms-topbar { margin: 10px 10px 0; }
    }
    @media (max-width: 767px) {
      .vessel-dms-topbar { margin: 8px 8px 0; height: 60px !important; border-radius: 18px; }
      .vessel-dms-app h1 { font-size: 22px !important; }
      .vessel-dms-app h2 { font-size: 19px !important; }
      .vessel-dms-app input[style*="min-width"], .vessel-dms-app select[style*="min-width"], .vessel-dms-app textarea[style*="min-width"] { min-width: 0 !important; }
      .vessel-dms-app input[style*="width: 2"], .vessel-dms-app input[style*="width: 1"], .vessel-dms-app select[style*="width: 1"], .vessel-dms-app select[style*="width: 2"] { width: 100% !important; }
      .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] > div:only-child { border-radius: 16px; }
    }
    @media (max-height: 480px) and (orientation: landscape) {
      .vessel-dms-topbar { height: 52px !important; margin-top: 6px; }
    }

    /* Reduced motion: only the ambient background loops calm down. The
       sign-in voyage and deck doors are short, user-triggered transitions and
       still play (set localStorage 'vdmsReduceMotion' = '1' to skip them too). */
    @media (prefers-reduced-motion: reduce) {
      .vessel-dms-app::before, .vessel-dms-auth::before, .vdms-auth-photo, .vdms-bob, .vdms-sheen3d, .vdms-shimmer, .vdms-pulse {
        animation: none !important;
      }
    }
  `;
}
