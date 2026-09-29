// Phase 6 — "Ocean Clay" visual language (chosen 2026-09-21 from the three
// claymorphism options presented on the design canvas: Soft Pastel Clay /
// Warm Sand Clay / Ocean Clay). Teal/navy palette themed to the vessel
// domain, applied as a soft "puffy" claymorphism style: dual-tone shadows
// (a dark cast shadow + a light highlight) and generous corner radii,
// rather than flat cards/buttons.
//
// This is a first pass covering the shared, highest-visual-impact pieces
// (Folder View's tile cards, the new Add Folder dialog, primary vessel
// actions). It is NOT yet applied line-by-line across every one of
// VesselsPage.tsx / DocumentsPage.tsx / VesselFormModal.tsx's inline
// styles — those files are 1,000-4,500+ lines each of hand-written inline
// styles with no shared Button/Card/Badge component to retheme centrally,
// so a full pass is a larger follow-up. Import these tokens wherever the
// next pass extends the look, so the palette stays consistent instead of
// re-deriving colors per file.
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
  bg: '#f8f1ea',
  surface: '#fff9f5',
  surfaceRaised: '#f7e6d8',
  surfaceHover: '#f2e2d5',

  text: '#342417',
  textMuted: '#8a6552',

  accent: '#DD9159',
  accentHover: '#C77A3E',
  accentDark: '#B96E35',
  accentDeep: '#995B2D',
  accentSoft: '#f7d8bf',
  accentSoftHover: '#f0cab1',

  accentGradient: 'linear-gradient(150deg, #E7A66D, #DD9159)',
  accentGradientHover: 'linear-gradient(150deg, #D38D56, #C77A3E)',
  iconBgGradient: 'linear-gradient(150deg, #f9dcc0, #efb57a)',

  // Card / tile: soft dual-tone "puffy" shadow — dark cast + light highlight.
  shadowRaised: '8px 8px 18px rgba(221,145,89,0.22), -8px -8px 16px rgba(255,255,255,0.88)',
  shadowRaisedHover: '10px 10px 22px rgba(221,145,89,0.28), -10px -10px 20px rgba(255,255,255,0.92)',

  // Primary button: gradient fill + outer cast shadow + inset highlight/shadow for a puffy 3D look.
  shadowButton: '0 10px 22px rgba(221,145,89,0.35), inset 0 2px 3px rgba(255,255,255,0.45), inset 0 -3px 6px rgba(150,89,42,0.28)',

  // Icon badge: inset highlight (top) + inset shadow (bottom).
  shadowIcon: 'inset 0 2px 3px rgba(255,255,255,0.7), inset 0 -3px 5px rgba(185,110,53,0.22)',

  // Status pills.
  pillActiveBg: '#cdeedb',
  pillActiveText: '#245a3d',
  pillActiveShadow: 'inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -2px 3px rgba(36,90,61,0.18)',
  pillWarnBg: '#e3d9c2',
  pillWarnText: '#7a6420',
  pillWarnShadow: 'inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -2px 3px rgba(122,100,32,0.18)',
  pillDangerBg: '#ecccc8',
  pillDangerText: '#8a3226',
  pillDangerShadow: 'inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -2px 3px rgba(138,50,38,0.18)',

  radiusCard: 24,
  radiusTile: 22,
  radiusButton: 20,
  radiusIcon: 14,
} as const;

// Theme-aware tokens. Each value is a CSS custom property that flips between
// light and night via the [data-vessel-theme] blocks in AppLayout.tsx, with the
// light value as fallback (so content portaled outside the app root, e.g.
// Fluent Layers, still renders). Accent/background/text tokens are also the
// ones Settings → Color Management rewrites at runtime (applyColorTheme, below).
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
  // Pre-mixed accent-tinted shadow color for glows that used to be built by
  // string-concatenating an alpha suffix onto a raw hex accent (e.g.
  // `${clay.accent}88`) — that trick breaks now that clay.accent is a
  // var() reference, so call sites use this token instead.
  accentGlow: v('accent-glow', 'rgba(221,145,89,0.53)'),
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

/** The 4 colors an admin can set, per mode. Everything else the theme needs
 * (surfaces, soft accent tints, shadows, Fluent's palette) is derived from
 * these via the helpers below, so the settings UI stays to 4 pickers per
 * mode instead of asking someone to hand-tune 20 hex values. */
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
  night: { bg: '#211812', text: '#f8eee6', accent: clayLight.accent, hover: clayLight.accentHover },
};

/** Ready-made palettes offered next to the custom pickers in Color
 * Management, in addition to the shipped Ocean Clay default. Purely a
 * starting point — every value stays editable afterwards. Must stay in
 * sync with PRESETS in backend/app/color_settings_api.py. */
export const CLAY_COLOR_PRESETS: { id: string; label: string; colors: ClayColorTheme }[] = [
  { id: 'ocean-clay', label: 'Ocean Clay (Default)', colors: DEFAULT_CLAY_COLORS },
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

/** Derives every other token the claymorphism look needs from the 4 colors
 * an admin actually sets. Approximate by design — matched to the original
 * Ocean Clay ratios closely enough to look coherent for any accent choice,
 * not pixel-identical to the hand-picked defaults. */
function deriveTokens(c: ClayColorSet, isNight: boolean) {
  const accentSoft = isNight ? mixHex(c.accent, '#000000', 0.62) : mixHex(c.accent, '#ffffff', 0.78);
  const accentSoftHover = isNight ? mixHex(c.accent, '#000000', 0.54) : mixHex(c.accent, '#ffffff', 0.70);
  const accentDark = darkenHex(c.hover, 0.08);
  const accentDeep = darkenHex(c.hover, 0.22);
  const surface = lightenHex(c.bg, 0.06);
  const surfaceRaised = isNight ? lightenHex(c.bg, 0.12) : mixHex(c.bg, c.accent, 0.06);
  const surfaceHover = isNight ? lightenHex(c.bg, 0.18) : mixHex(c.bg, c.accent, 0.10);
  const textMuted = mixHex(c.text, c.bg, 0.45);
  return { accentSoft, accentSoftHover, accentDark, accentDeep, surface, surfaceRaised, surfaceHover, textMuted };
}

/** Rewrites the `--clay-*` custom properties for both [data-vessel-theme]
 * blocks so the whole app (everything importing `clay` from this file)
 * retheme instantly — used both for the saved app-wide colors (loaded once
 * on mount) and for Color Management's live preview while editing, before
 * Save is clicked. `!important` so this always wins over the static
 * defaults in AppLayout.tsx's injectFullScreenStyles, regardless of which
 * <style> tag landed in <head> first. */
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
      --clay-accent-gradient: linear-gradient(150deg, ${lightenHex(c.accent, 0.12)}, ${c.accent}) !important;
      --clay-accent-gradient-hover: linear-gradient(150deg, ${c.accent}, ${c.hover}) !important;
      --clay-icon-bg: linear-gradient(150deg, ${d.accentSoft}, ${d.accentSoftHover}) !important;
      --clay-accent-glow: ${hexToRgba(c.accent, 0.53)} !important;
      --clay-shadow-button: 0 10px 22px ${hexToRgba(c.accent, 0.35)}, inset 0 2px 3px rgba(255,255,255,0.45), inset 0 -3px 6px ${hexToRgba(d.accentDeep, 0.28)} !important;
    `;
  };

  style.textContent = `
    [data-vessel-theme="light"] { ${block(theme.light, false)} }
    [data-vessel-theme="night"] { ${block(theme.night, true)} }
  `;
}

/** Builds the Fluent UI theme (dropdowns, spinners, the handful of stock
 * Fluent controls still in use) from the same 4 admin-set colors, instead
 * of the old hand-picked static palette. Approximate derivation — see
 * deriveTokens() note above. */
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
    fonts: {
      medium: { fontSize: '15px' },
      mediumPlus: { fontSize: '16px' },
      large: { fontSize: '20px' },
      xLarge: { fontSize: '24px' },
    },
  };
}

// Static exports kept for any call site that hasn't moved to
// buildDeepHarborTheme(host.state.colorTheme.<mode>, isNight) yet — equal to
// buildDeepHarborTheme(DEFAULT_CLAY_COLORS.<mode>, ...), i.e. the original
// Ocean Clay look.
export const deepHarborTheme: IPartialTheme = buildDeepHarborTheme(DEFAULT_CLAY_COLORS.light, false);
export const deepHarborNightTheme: IPartialTheme = buildDeepHarborTheme(DEFAULT_CLAY_COLORS.night, true);
