// Maritime futuristic skin — DESIGN ONLY.
// A single extra <style> tag layered on top of AppLayout.tsx's
// injectFullScreenStyles(). It only re-maps CSS custom properties and adds
// presentation rules (glass, radii, fonts, background, motion). No handler,
// state, API call, permission or DOM order is touched.
//
// Settings → Color Management keeps working: every --clay-* value that
// applyColorTheme() writes (bg / text / accent / hover and their derived
// surfaces, gradients, glow) is deliberately NOT overridden here.

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
    /* ── Neutral + surface tokens (not admin-controlled) ───────────────── */
    html [data-vessel-theme="light"] {
      --vdms-surface: #ffffff;
      --vdms-surface-alt: #f6f8fb;
      --vdms-border: #e2e8f0;
      --vdms-border-soft: #eef1f5;
      --vdms-text: #101b2d;
      --vdms-text-secondary: #475569;
      --vdms-text-muted: #64748b;
      --vdms-text-faint: #94a3b8;
      --vdms-toggle-active-bg: #0b2a4a;
      --vdms-toggle-active-text: #ffffff;
      --vdms-glass: #ffffff;
      --vdms-glass-strong: #ffffff;
      --vdms-line: #e2e8f0;
      --vdms-line-strong: #cbd5e1;
      --vdms-field: #ffffff;
      --vdms-shadow: 0 1px 2px rgba(16,27,45,0.04);
      --vdms-focus: rgba(14,116,144,0.6);
      --vdms-focus-soft: rgba(14,116,144,0.14);
      --clay-shadow-raised: 0 1px 2px rgba(16,27,45,0.05), 0 4px 12px rgba(16,27,45,0.06);
      --clay-shadow-raised-hover: 0 2px 4px rgba(16,27,45,0.06), 0 8px 20px rgba(16,27,45,0.10);
      --clay-shadow-icon: inset 0 0 0 1px rgba(14,116,144,0.14);
      --clay-pill-active-bg: #dcfce7; --clay-pill-active-text: #15803d; --clay-pill-active-shadow: none;
      --clay-pill-warn-bg: #fef3c7;   --clay-pill-warn-text: #b45309;   --clay-pill-warn-shadow: none;
      --clay-pill-danger-bg: #fee2e2; --clay-pill-danger-text: #b91c1c; --clay-pill-danger-shadow: none;
    }
    html [data-vessel-theme="night"] {
      --vdms-surface: #101d2e;
      --vdms-surface-alt: #152536;
      --vdms-border: rgba(148,178,204,0.18);
      --vdms-border-soft: rgba(148,178,204,0.1);
      --vdms-text: #e7eef5;
      --vdms-text-secondary: #aebfd1;
      --vdms-text-muted: #8ca0b5;
      --vdms-text-faint: #71869b;
      --vdms-toggle-active-bg: #2dd4bf;
      --vdms-toggle-active-text: #04211d;
      --vdms-glass: #101d2e;
      --vdms-glass-strong: #0a1626;
      --vdms-line: rgba(148,178,204,0.18);
      --vdms-line-strong: rgba(148,178,204,0.3);
      --vdms-field: #101d2e;
      --vdms-shadow: 0 1px 2px rgba(0,0,0,0.3);
      --vdms-focus: rgba(45,212,191,0.65);
      --vdms-focus-soft: rgba(45,212,191,0.16);
      --clay-shadow-raised: 0 1px 2px rgba(0,0,0,0.3), 0 4px 14px rgba(0,0,0,0.35);
      --clay-shadow-raised-hover: 0 2px 4px rgba(0,0,0,0.35), 0 8px 22px rgba(0,0,0,0.4);
      --clay-shadow-icon: inset 0 0 0 1px rgba(45,212,191,0.18);
      --clay-pill-active-bg: rgba(34,197,94,0.16); --clay-pill-active-text: #86efac; --clay-pill-active-shadow: none;
      --clay-pill-warn-bg: rgba(245,158,11,0.16);  --clay-pill-warn-text: #fcd34d;  --clay-pill-warn-shadow: none;
      --clay-pill-danger-bg: rgba(239,68,68,0.16); --clay-pill-danger-text: #fca5a5; --clay-pill-danger-shadow: none;
    }

    /* ── Type ─────────────────────────────────────────────────────────── */
    .vessel-dms-app, .vessel-dms-auth { font-family: ${vdmsFont.ui} !important; -webkit-font-smoothing: antialiased; }
    .vessel-dms-app h1, .vessel-dms-app h2, .vessel-dms-app h3, .vessel-dms-auth h1 { font-family: ${vdmsFont.display}; letter-spacing: -0.02em; }

    /* ── App backdrop ─────────────────────────────────────────────────── */
    /* Flat, solid enterprise background — no photographic backdrop. */
    .vessel-dms-app[data-vessel-theme="light"] { background: var(--clay-bg, #eef2f7) !important; }
    .vessel-dms-app[data-vessel-theme="night"] { background: var(--clay-bg, #0a1626) !important; }
    .vessel-dms-auth[data-vessel-theme="light"] {
      background: linear-gradient(160deg, #0b2a4a 0%, #0e3a63 55%, #0a4a5c 100%) !important;
    }
    .vessel-dms-auth[data-vessel-theme="night"] {
      background: linear-gradient(160deg, #061021 0%, #0a1626 55%, #0a2320 100%) !important;
    }
    .vessel-dms-app { isolation: isolate; }

    /* ── Sidebar shell ─────────────────────────────────────────────────── */
    /* Background colour itself comes from the inline clay.accentDeep style
       in Sidebar.tsx (so it stays admin-preset aware). Full-bleed enterprise
       nav rail — flush to the viewport edges, no floating-card margin/radius
       (previously margin: 14px 0 14px 14px + border-radius: 16px, which left
       a visible gap of page background above/left of the rail). */
    .vessel-dms-sidebar {
      margin: 0; height: 100vh !important;
      transition: width .25s ${vdmsMotion.glide}, min-width .25s ${vdmsMotion.glide}, max-width .25s ${vdmsMotion.glide}, transform .25s ${vdmsMotion.glide}, opacity .2s ease !important;
    }
    .vessel-dms-sidebar-nav button { transition: background .2s ease, color .2s ease !important; }
    .vessel-dms-sidebar button { border-radius: 8px; font-size: 15px; }
    @media (max-width: 1024px) { .vessel-dms-sidebar { top: 0 !important; left: 0 !important; margin: 0; } }

    /* ── Top bar ──────────────────────────────────────────────────────── */
    /* Flush against the sidebar and the viewport top — no floating-card
       margin/radius (previously margin: 14px 14px 0 + border-radius: 12px,
       which left a visible gap between the sidebar and the header). */
    .vessel-dms-topbar {
      margin: 0; height: 72px !important; border-radius: 0;
      position: relative; z-index: 5;
    }
    .vessel-dms-topbar, .vessel-dms-topbar button { font-family: ${vdmsFont.ui}; }

    /* ── Dashboard hero + glass cards ─────────────────────────────────── */
    html [data-vessel-theme="light"] .vdms-hero-ovl { background: linear-gradient(90deg, rgba(236,247,255,0.95) 0%, rgba(236,247,255,0.7) 55%, rgba(236,247,255,0.05) 100%); }
    html [data-vessel-theme="night"] .vdms-hero-ovl { background: linear-gradient(90deg, rgba(3,20,35,0.92) 0%, rgba(3,20,35,0.6) 55%, rgba(3,20,35,0.15) 100%); }
    html [data-vessel-theme="night"] .vdms-hero { background-color: #0a1626; filter: none; }
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

    /* ── Popups: subtle blurred overlay + soft entrance ───────────────── */
    .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] {
      backdrop-filter: blur(3px); -webkit-backdrop-filter: blur(3px);
      animation: vdms-fade .2s ease both;
    }
    .vessel-dms-app div[style*="position: fixed"][style*="inset: 0px"] > div:only-child {
      border-radius: 14px; animation: vdms-pop .25s ${vdmsMotion.glide} both;
    }

    /* ── Motion primitives: restrained, functional transitions only ───── */
    @keyframes vdms-fade { from { opacity: 0 } to { opacity: 1 } }
    @keyframes vdms-pop { from { opacity: 0; transform: translateY(8px) scale(.98) } to { opacity: 1; transform: none } }
    @keyframes vdms-pulse { 0%, 100% { box-shadow: 0 0 0 0 rgba(16,185,129,0.45) } 50% { box-shadow: 0 0 0 6px rgba(16,185,129,0) } }
    .vdms-pop { animation: vdms-pop .3s ${vdmsMotion.glide} both; }
    .vdms-pulse { animation: vdms-pulse 2.4s ease-in-out infinite; }

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
    html .vessel-dms-app .vessel-dms-sidebar { height: calc(100vh / var(--vdms-z)) !important; }
    @supports (height: 100dvh) { html .vessel-dms-app .vessel-dms-sidebar { height: calc(100dvh / var(--vdms-z)) !important; } }
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

    @media (prefers-reduced-motion: reduce) {
      .vdms-pulse, .vdms-pop { animation: none !important; }
    }
  `;
}
