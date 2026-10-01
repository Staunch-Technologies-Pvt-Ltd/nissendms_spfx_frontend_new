
// Visual refresh patch — DESIGN ONLY. One extra <style> tag injected AFTER
// injectFuturisticTheme(). It only overrides presentation; no handler, state,
// API call or DOM order is touched.
//
//  1. Top-bar breadcrumb ("Vessel DMS ›") was glass-on-glass (unreadable) → solid ink.
//  2. Table text was too pale on the glass → stronger ink for cells and headers.
//  3. Hero banner photo was buried under a 95% wash → photo reads clearly again.
//  4. Stray orange dot: the hover/focus ring was drawn on zero-size elements → skipped.

const STYLE_ID = 'vessel-dms-refresh';

export function injectRefreshTheme(): void {
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  document.head.appendChild(style);
  style.textContent = `
    /* 1 · breadcrumb contrast */
    html [data-vessel-theme="light"] .vessel-dms-topbar .vdms-crumb { color: #0b3a66 !important; text-shadow: 0 1px 0 rgba(255,255,255,0.6); }
    html [data-vessel-theme="night"] .vessel-dms-topbar .vdms-crumb { color: #d9eefb !important; }

    /* 2 · table legibility */
    html [data-vessel-theme="light"] .vessel-dms-app table td { color: #0a2a44; }
    html [data-vessel-theme="light"] .vessel-dms-app table th { color: #2c4a63 !important; font-weight: 700; }
    html [data-vessel-theme="night"] .vessel-dms-app table td { color: #eaf6fd; }
    html [data-vessel-theme="night"] .vessel-dms-app table th { color: #b8d2e3 !important; }
    html [data-vessel-theme="light"] .vessel-dms-app table td span[style*="opacity"] { opacity: 1 !important; }

    /* 3 · hero banner: let the ship photo show */
    html [data-vessel-theme="light"] .vdms-hero-ovl { background: linear-gradient(90deg, rgba(236,247,255,0.92) 0%, rgba(236,247,255,0.55) 38%, rgba(236,247,255,0) 72%); }
    html [data-vessel-theme="night"] .vdms-hero-ovl { background: linear-gradient(90deg, rgba(3,20,35,0.9) 0%, rgba(3,20,35,0.5) 38%, rgba(3,20,35,0) 72%); }
    .vdms-hero { filter: saturate(1.25) contrast(1.06); }

    /* 4 · no focus ring on empty / zero-size elements */
    html .vessel-dms-app button:empty:hover, html .vessel-dms-app a:empty:hover, html .vessel-dms-app [role="button"]:empty:hover,
    html .vessel-dms-app span:empty:hover, html .vessel-dms-app td:empty:hover {
      outline: none !important; box-shadow: none !important; filter: none !important;
    }

    /* 5 · Settings module — full-width glass page, modern tables, bigger type */
    .vdms-settings-body { font-size: 16px; color: var(--vdms-text); }
    .vdms-settings-body h3, .vdms-settings-body h4 { font-family: 'Sora','Segoe UI Variable',sans-serif; letter-spacing: -0.02em; color: var(--vdms-text) !important; }
    .vdms-settings-body h4 { font-size: 19px !important; }
    .vdms-settings-body p, .vdms-settings-body label { font-size: 16px; }
    /* hard-coded slate text -> app ink (both modes) */
    .vdms-settings [style*="color: rgb(15, 23, 42)"], .vdms-settings [style*="color: rgb(30, 41, 59)"], .vdms-settings [style*="color: rgb(51, 65, 85)"] { color: var(--vdms-text) !important; }
    .vdms-settings [style*="color: rgb(71, 85, 105)"], .vdms-settings [style*="color: rgb(100, 116, 139)"] { color: var(--vdms-text-secondary) !important; }
    .vdms-settings [style*="color: rgb(148, 163, 184)"], .vdms-settings [style*="color: rgb(156, 163, 175)"] { color: var(--vdms-text-muted) !important; }
    /* hard-coded white / slate panels -> glass */
    .vdms-settings-body [style*="background: rgb(255, 255, 255)"], .vdms-settings-body [style*="background-color: rgb(255, 255, 255)"] { background: var(--vdms-field) !important; }
    .vdms-settings-body [style*="background: rgb(248, 250, 252)"], .vdms-settings-body [style*="background: rgb(241, 245, 249)"],
    .vdms-settings-body [style*="background: rgb(240, 253, 244)"], .vdms-settings-body [style*="background: rgb(236, 253, 245)"],
    .vdms-settings-body [style*="background: rgb(240, 249, 255)"], .vdms-settings-body [style*="background: rgb(239, 246, 255)"] { background: var(--vdms-surface-alt) !important; }
    .vdms-settings-body [style*="rgb(226, 232, 240)"], .vdms-settings-body [style*="rgb(241, 245, 249)"], .vdms-settings-body [style*="rgb(203, 213, 225)"] { border-color: var(--vdms-line) !important; }
    .vdms-settings-body [style*="border-radius: 6px"], .vdms-settings-body [style*="border-radius: 8px"] { border-radius: 14px !important; }
    .vdms-settings-body [style*="border-radius: 10px"], .vdms-settings-body [style*="border-radius: 12px"] { border-radius: 18px !important; }
    /* controls */
    .vdms-settings-body input:not([type="checkbox"]):not([type="radio"]):not([type="color"]), .vdms-settings-body select, .vdms-settings-body textarea { font-size: 16px !important; min-height: 46px; padding: 8px 14px !important; border-radius: 14px !important; border: 1px solid var(--vdms-line-strong) !important; background: var(--vdms-field) !important; color: var(--vdms-text) !important; box-sizing: border-box; }
    .vdms-settings-body button { font-family: 'Manrope','Segoe UI Variable',sans-serif; font-size: 14px !important; font-weight: 700; border-radius: 12px; min-height: 36px; }
    /* modern tables */
    .vdms-settings-body table { width: 100%; border-collapse: separate !important; border-spacing: 0; border: 1px solid var(--vdms-line); border-radius: 20px; overflow: hidden; background: var(--vdms-field); font-size: 16px !important; }
    .vdms-settings-body table thead tr { border: none !important; background: var(--vdms-surface-alt) !important; }
    .vdms-settings-body table th { padding: 16px 18px !important; font-family: 'JetBrains Mono',Consolas,monospace !important; font-size: 12px !important; font-weight: 700 !important; letter-spacing: 0.12em !important; text-transform: uppercase; text-align: left; color: var(--vdms-text-muted) !important; background: var(--vdms-surface-alt) !important; border-bottom: 1px solid var(--vdms-line-strong) !important; white-space: nowrap; }
    .vdms-settings-body table td { padding: 16px 18px !important; font-size: 16px !important; color: var(--vdms-text) !important; border-bottom: 1px solid var(--vdms-line) !important; vertical-align: middle; }
    .vdms-settings-body table tbody tr { border: none !important; transition: background .2s ease; }
    .vdms-settings-body table tbody tr:nth-child(even) td { background: rgba(10,126,168,0.045); }
    .vdms-settings-body table tbody tr:hover td { background: var(--vdms-focus-soft); }
    .vdms-settings-body table tbody tr:last-child td { border-bottom: none !important; }
    .vdms-settings-body td [style*="font-size: 11px"], .vdms-settings-body td [style*="font-size: 12px"] { font-size: 13px !important; }
  `;
}
