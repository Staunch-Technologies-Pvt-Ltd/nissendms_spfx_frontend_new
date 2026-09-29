// Shared inline-style tokens for the Migration Assistant module. SPFx has no
// Tailwind/CSS-in-JS pipeline (see clayTheme.ts's header comment on the same
// constraint), so this mirrors that file's approach: plain TS style objects,
// themed off the host app's existing clay palette so the module looks native
// rather than bolted on.
import type { CSSProperties } from 'react';
import { clay } from '../clayTheme';

export function tokens(isNight: boolean) {
  const surface = isNight ? '#2b211b' : '#ffffff';
  const surfaceAlt = isNight ? '#3a291f' : '#f8fafc';
  const border = isNight ? '#493225' : '#e2e8f0';
  const text = isNight ? '#f8eee6' : '#0f172a';
  const textMuted = isNight ? '#c7a58d' : '#64748b';
  const textSubtle = isNight ? '#a8886f' : '#94a3b8';
  return { surface, surfaceAlt, border, text, textMuted, textSubtle };
}

export const success = '#16a34a';
export const warning = '#d97706';
export const error = '#dc2626';

export function card(isNight: boolean): CSSProperties {
  const t = tokens(isNight);
  return { borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface };
}

export function primaryBtn(disabled?: boolean): CSSProperties {
  return {
    display: 'inline-flex', alignItems: 'center', gap: 6,
    background: disabled ? `${clay.accent}80` : clay.accentGradient,
    color: '#fff', border: 'none', borderRadius: 10, padding: '8px 16px',
    fontSize: 13, fontWeight: 600, cursor: disabled ? 'not-allowed' : 'pointer',
    boxShadow: disabled ? 'none' : clay.shadowButton,
  };
}

export function secondaryBtn(isNight: boolean): CSSProperties {
  const t = tokens(isNight);
  return {
    display: 'inline-flex', alignItems: 'center', gap: 6,
    background: t.surfaceAlt, color: t.textMuted, border: `1px solid ${t.border}`,
    borderRadius: 10, padding: '8px 14px', fontSize: 13, fontWeight: 500, cursor: 'pointer',
  };
}

export function pill(kind: 'success' | 'warning' | 'error' | 'primary'): CSSProperties {
  const map: Record<string, [string, string]> = {
    success: ['#dcfce7', success],
    warning: ['#fef3c7', warning],
    error: ['#fee2e2', error],
    primary: [`${clay.accent}22`, clay.accentDeep],
  };
  const [bg, fg] = map[kind];
  return { display: 'inline-block', borderRadius: 999, padding: '2px 9px', fontSize: 11, fontWeight: 600, background: bg, color: fg };
}
