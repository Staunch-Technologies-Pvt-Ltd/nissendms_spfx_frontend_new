// Phase 8 — "Enterprise Maritime" visual language (deep navy + teal accent,
// flat solid surfaces, restrained shadows). Replaces the Phase 7 "Maritime"
// glass/gradient defaults. DESIGN ONLY: every export, key and function
// signature below is unchanged, so nothing that imports `clay` needs to
// change. Ocean Clay is kept as a selectable preset in Settings → Color
// Management ("Ocean Clay (Legacy)").
//
// React 17 / Fluent UI 8 / SPFx constraint: plain TS constants + inline
// style objects only — no CSS-in-JS library, no Tailwind, no build changes.
//
// Settings → Color Management (color_settings_api.py) lets an admin
// recolor background / text / design(accent) / hover for light and night
// mode independently. Every value below that participates in that
// recolor is CSS-variable-backed (`v('name', fallback)`) so applyColorTheme()
// can retheme the whole app at runtime just by rewriting those custom
// properties — nothing importing `clay` needs to change.

import type { IPartialTheme } from '@fluentui/react';

// Raw light-mode values. Used for Fluent UI theme palettes (which need real
// colours) and as the fallback inside every var() below.
export const clayLight = {
  bg: '#eef2f7',
  surface: '#ffffff',
  surfaceRaised: '#f6f8fb',
  surfaceHover: '#eaf0f6',

  text: '#101b2d',
  textMuted: '#5b6b7f',

  accent: '#0e7490',
  accentHover: '#0b5c72',
  accentDark: '#0a4a5c',
  accentDeep: '#0b2a4a',
  accentSoft: '#dceef2',
  accentSoftHover: '#c7e4ea',

  accentGradient: 'linear-gradient(135deg, #14919b, #0e7490)',
  accentGradientHover: 'linear-gradient(135deg, #0e7490, #0b5c72)',
  iconBgGradient: 'linear-gradient(150deg, #e3f2f4, #cfe8ec)',

  // Card / tile: flat surface, soft low-opacity lift.
  shadowRaised: '0 1px 2px rgba(16,27,45,0.05), 0 4px 12px rgba(16,27,45,0.06)',
  shadowRaisedHover: '0 2px 4px rgba(16,27,45,0.06), 0 8px 20px rgba(16,27,45,0.10)',

  // Primary button: flat fill + a restrained accent-tinted shadow.
  shadowButton: '0 1px 2px rgba(16,27,45,0.08), 0 4px 10px rgba(14,116,144,0.25)',

  // Icon badge.
  shadowIcon: 'inset 0 0 0 1px rgba(14,116,144,0.14)',

  // Status pills — green/amber/red semantic states.
  pillActiveBg: '#dcfce7',
  pillActiveText: '#15803d',
  pillActiveShadow: 'none',
  pillWarnBg: '#fef3c7',
  pillWarnText: '#b45309',
  pillWarnShadow: 'none',
  pillDangerBg: '#fee2e2',
  pillDangerText: '#b91c1c',
  pillDangerShadow: 'none',

  radiusCard: 12,
  radiusTile: 10,
  radiusButton: 8,
  radiusIcon: 8,
} as const;

// Theme-aware tokens. Each value is a CSS custom property that flips between
// light and night via the [data-vessel-theme] blocks in AppLayout.tsx /
// futuristicTheme.ts, with the light value as fallback.
const v = (name: string, fallback: string): string => `var(--clay-${name}, ${fallback})`;

export const clay = {
  bg: v('bg', clayLight.bg),
  surface: v('surface', clayLight.surface),
  surfaceRaised: v('surface-raised', clayLight.surfaceRaised),
  surfaceHover: v('surface-hover', clayLight.surfaceHover),
  text: v('text', clayLight.text),
  textMuted: v('text-muted', clayLight.textMuted),
  accent: v('accent', clayLight.accent),
  accentHover: v('accent-hover', clayLight.accentHover),
  accentDark: v('accent-dark', clayLight.accentDark),
  accentDeep: v('accent-deep', clayLight.accentDeep),
  accentSoft: v('accent-soft', clayLight.accentSoft),
  accentSoftHover: v('accent-soft-hover', clayLight.accentSoftHover),
  accentGlow: v('accent-glow', 'rgba(10,126,168,0.4)'),
  accentGradient: v('accent-gradient', clayLight.accentGradient),
  accentGradientHover: v('accent-gradient-hover', clayLight.accentGradientHover),
  iconBgGradient: v('icon-bg', clayLight.iconBgGradient),
  shadowRaised: v('shadow-raised', clayLight.shadowRaised),
  shadowRaisedHover: v('shadow-raised-hover', clayLight.shadowRaisedHover),
  shadowButton: v('shadow-button', clayLight.shadowButton),
  shadowIcon: v('shadow-icon', clayLight.shadowIcon),
  pillActiveBg: v('pill-active-bg', clayLight.pillActiveBg),
  pillActiveText: v('pill-active-text', clayLight.pillActiveText),
  pillActiveShadow: v('pill-active-shadow', clayLight.pillActiveShadow),
  pillWarnBg: v('pill-warn-bg', clayLight.pillWarnBg),
  pillWarnText: v('pill-warn-text', clayLight.pillWarnText),
  pillWarnShadow: v('pill-warn-shadow', clayLight.pillWarnShadow),
  pillDangerBg: v('pill-danger-bg', clayLight.pillDangerBg),
  pillDangerText: v('pill-danger-text', clayLight.pillDangerText),
  pillDangerShadow: v('pill-danger-shadow', clayLight.pillDangerShadow),
  radiusCard: clayLight.radiusCard,
  radiusTile: clayLight.radiusTile,
  radiusButton: clayLight.radiusButton,
  radiusIcon: clayLight.radiusIcon,
};

// ─────────────────────────────────────────────────────────────────────────
// Settings → Color Management
// ─────────────────────────────────────────────────────────────────────────

export interface ClayColorSet {
  bg: string;
  text: string;
  accent: string;
  hover: string;
}

export interface ClayColorTheme {
  light: ClayColorSet;
  night: ClayColorSet;
}

// Must stay in sync with DEFAULT_COLORS in backend/app/color_settings_api.py.
export const DEFAULT_CLAY_COLORS: ClayColorTheme = {
  light: { bg: clayLight.bg, text: clayLight.text, accent: clayLight.accent, hover: clayLight.accentHover },
  night: { bg: '#0a1626', text: '#e7eef5', accent: '#2dd4bf', hover: '#14b8a6' },
};

/** Must stay in sync with PRESETS in backend/app/color_settings_api.py. */
export const CLAY_COLOR_PRESETS: { id: string; label: string; colors: ClayColorTheme }[] = [
  { id: 'maritime', label: 'Maritime (Default)', colors: DEFAULT_CLAY_COLORS },
  {
    id: 'ocean-clay', label: 'Ocean Clay (Legacy)', colors: {
      light: { bg: '#f8f1ea', text: '#342417', accent: '#DD9159', hover: '#C77A3E' },
      night: { bg: '#211812', text: '#f8eee6', accent: '#DD9159', hover: '#C77A3E' },
    },
  },
  {
    id: 'slate-blue', label: 'Slate Blue', colors: {
      light: { bg: '#eef2f8', text: '#1e293b', accent: '#3b6fd6', hover: '#2f59b0' },
      night: { bg: '#141b29', text: '#e7edf7', accent: '#4f83e6', hover: '#3b6fd6' },
    },
  },
  {
    id: 'forest-sage', label: 'Forest Sage', colors: {
      light: { bg: '#f0f4ec', text: '#26331f', accent: '#5c8a4a', hover: '#4a7239' },
      night: { bg: '#182016', text: '#e7f0e0', accent: '#6ea057', hover: '#5c8a4a' },
    },
  },
  {
    id: 'sunset-coral', label: 'Sunset Coral', colors: {
      light: { bg: '#fcf1ec', text: '#3a2118', accent: '#e2634a', hover: '#c94f38' },
      night: { bg: '#241713', text: '#fbe9e2', accent: '#e97a63', hover: '#e2634a' },
    },
  },
  {
    id: 'midnight-indigo', label: 'Midnight Indigo', colors: {
      light: { bg: '#eeeefb', text: '#221f3d', accent: '#5b4bd6', hover: '#4636b5' },
      night: { bg: '#161327', text: '#e8e6fb', accent: '#7566e6', hover: '#5b4bd6' },
    },
  },
];

// ── Tiny color-math helpers (no dependency — plain hex math) ───────────────
function clampByte(n: number): number { return Math.max(0, Math.min(255, Math.round(n))); }

function hexToRgb(hex: string): [number, number, number] {
  const h = (hex || '').replace('#', '').trim();
  const full = h.length === 3 ? h.split('').map(c => c + c).join('') : h.padEnd(6, '0').slice(0, 6);
  const int = parseInt(full, 16);
  const safe = Number.isNaN(int) ? 0 : int;
  return [(safe >> 16) & 255, (safe >> 8) & 255, safe & 255];
}

function rgbToHex(r: number, g: number, b: number): string {
  return '#' + [r, g, b].map(x => clampByte(x).toString(16).padStart(2, '0')).join('');
}

/** Mix hexA toward hexB by t (0 = hexA, 1 = hexB). */
export function mixHex(hexA: string, hexB: string, t: number): string {
  const [r1, g1, b1] = hexToRgb(hexA);
  const [r2, g2, b2] = hexToRgb(hexB);
  return rgbToHex(r1 + (r2 - r1) * t, g1 + (g2 - g1) * t, b1 + (b2 - b1) * t);
}

export function lightenHex(hex: string, t: number): string { return mixHex(hex, '#ffffff', t); }
export function darkenHex(hex: string, t: number): string { return mixHex(hex, '#000000', t); }

export function hexToRgba(hex: string, alpha: number): string {
  const [r, g, b] = hexToRgb(hex);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

/** Derives every other token from the 4 colors an admin actually sets. */
function deriveTokens(c: ClayColorSet, isNight: boolean) {
  const accentSoft = isNight ? mixHex(c.accent, '#000000', 0.62) : mixHex(c.accent, '#ffffff', 0.8);
  const accentSoftHover = isNight ? mixHex(c.accent, '#000000', 0.54) : mixHex(c.accent, '#ffffff', 0.72);
  const accentDark = darkenHex(c.hover, 0.08);
  // Deep enough to read as a navy/near-black "primary" surface (sidebar,
  // topbar) regardless of which accent hue a preset uses.
  const accentDeep = darkenHex(c.hover, 0.5);
  const surface = isNight ? lightenHex(c.bg, 0.06) : lightenHex(c.bg, 0.6);
  const surfaceRaised = isNight ? lightenHex(c.bg, 0.12) : mixHex(lightenHex(c.bg, 0.4), c.accent, 0.04);
  const surfaceHover = isNight ? lightenHex(c.bg, 0.18) : mixHex(c.bg, c.accent, 0.1);
  const textMuted = mixHex(c.text, c.bg, 0.4);
  return { accentSoft, accentSoftHover, accentDark, accentDeep, surface, surfaceRaised, surfaceHover, textMuted };
}

/** Rewrites the `--clay-*` custom properties for both [data-vessel-theme]
 * blocks so the whole app retheme instantly (saved colors + live preview). */
export function applyColorTheme(theme: ClayColorTheme): void {
  if (typeof document === 'undefined') return;
  const id = 'vessel-dms-color-theme';
  let style = document.getElementById(id) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement('style');
    style.id = id;
    document.head.appendChild(style);
  }

  const block = (c: ClayColorSet, isNight: boolean): string => {
    const d = deriveTokens(c, isNight);
    return `
      --clay-bg: ${c.bg} !important;
      --clay-text: ${c.text} !important;
      --clay-text-muted: ${d.textMuted} !important;
      --clay-surface: ${d.surface} !important;
      --clay-surface-raised: ${d.surfaceRaised} !important;
      --clay-surface-hover: ${d.surfaceHover} !important;
      --clay-accent: ${c.accent} !important;
      --clay-accent-hover: ${c.hover} !important;
      --clay-accent-dark: ${d.accentDark} !important;
      --clay-accent-deep: ${d.accentDeep} !important;
      --clay-accent-soft: ${d.accentSoft} !important;
      --clay-accent-soft-hover: ${d.accentSoftHover} !important;
      --clay-accent-gradient: linear-gradient(135deg, ${lightenHex(c.accent, 0.12)}, ${c.hover}) !important;
      --clay-accent-gradient-hover: linear-gradient(135deg, ${c.accent}, ${darkenHex(c.hover, 0.12)}) !important;
      --clay-icon-bg: linear-gradient(150deg, ${d.accentSoft}, ${d.accentSoftHover}) !important;
      --clay-accent-glow: ${hexToRgba(c.accent, 0.4)} !important;
      --clay-shadow-button: 0 1px 2px rgba(16,27,45,0.08), 0 4px 10px ${hexToRgba(c.accent, 0.25)} !important;
    `;
  };

  style.textContent = `
    [data-vessel-theme="light"] { ${block(theme.light, false)} }
    [data-vessel-theme="night"] { ${block(theme.night, true)} }
  `;
}

/** Builds the Fluent UI theme from the same 4 admin-set colors. */
export function buildDeepHarborTheme(colors: ClayColorSet, isNight: boolean): IPartialTheme {
  const { accent, hover, bg, text } = colors;
  const neutralBase = isNight ? lightenHex(bg, 0.10) : mixHex(bg, accent, 0.08);
  return {
    palette: {
      themePrimary: accent,
      themeDark: darkenHex(hover, 0.15),
      themeDarker: darkenHex(hover, 0.35),
      themeSecondary: lightenHex(accent, 0.35),
      themeLighterAlt: lightenHex(accent, 0.97),
      themeLighter: lightenHex(accent, 0.90),
      themeLight: lightenHex(accent, 0.78),
      themeTertiary: hover,
      themeDarkAlt: hover,
      white: isNight ? bg : lightenHex(bg, 0.06),
      neutralLighterAlt: neutralBase,
      neutralLighter: isNight ? lightenHex(neutralBase, 0.10) : darkenHex(neutralBase, 0.06),
      neutralLight: isNight ? lightenHex(neutralBase, 0.20) : darkenHex(neutralBase, 0.14),
      neutralQuaternaryAlt: isNight ? lightenHex(neutralBase, 0.28) : darkenHex(neutralBase, 0.20),
      neutralPrimaryAlt: mixHex(text, bg, 0.30),
      neutralPrimary: text,
      neutralDark: darkenHex(text, isNight ? 0 : 0.15),
      black: isNight ? '#ffffff' : darkenHex(text, 0.30),
    },
    semanticColors: {
      bodyBackground: isNight ? bg : lightenHex(bg, 0.06),
      bodyText: text,
      disabledBackground: mixHex(bg, text, 0.12),
      disabledText: mixHex(text, bg, 0.45),
      variantBorder: isNight ? mixHex(accent, '#000000', 0.55) : lightenHex(accent, 0.78),
      variantBorderHovered: accent,
    },
    defaultFontStyle: { fontFamily: "'Manrope', 'Segoe UI Variable', 'Segoe UI', sans-serif" },
    fonts: {
      medium: { fontSize: '15px' },
      mediumPlus: { fontSize: '16px' },
      large: { fontSize: '20px' },
      xLarge: { fontSize: '24px' },
    },
  };
}

export const deepHarborTheme: IPartialTheme = buildDeepHarborTheme(DEFAULT_CLAY_COLORS.light, false);
export const deepHarborNightTheme: IPartialTheme = buildDeepHarborTheme(DEFAULT_CLAY_COLORS.night, true);
