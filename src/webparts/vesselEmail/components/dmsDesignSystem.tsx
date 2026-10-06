import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IDropdownStyleProps, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { clay } from './clayTheme';
import { vdmsFont } from './futuristicTheme';

/**
 * Shared enterprise UI kit — every page in the app should build its chrome
 * (page header, cards, buttons, inputs, tables, empty/loading states) from
 * these instead of reinventing its own inline styles. Originally built for
 * Documents, promoted to a shared module so Dashboard's look (the declared
 * master UI reference) propagates to every other page without copy-paste
 * drift. Every colour comes from `clay.*` or `var(--vdms-*)`, so Settings →
 * Color Management and the Night/Light toggle keep re-theming all of it.
 * Purely presentational — none of these helpers own state or wiring;
 * callers keep their own handlers/data.
 */

export const DMS_FONT_DISPLAY = "'Sora', 'Segoe UI Variable', 'Segoe UI', sans-serif";
// Text on a solid accent fill: the app's own toggle-active text token
// (white in light mode, deep teal-black in night mode, where the default
// accent is a light teal), so primary buttons stay readable in both themes.
export const DMS_ON_ACCENT = 'var(--vdms-toggle-active-text, #ffffff)';

export type DmsTone = 'accent' | 'success' | 'warning' | 'danger' | 'neutral';
export function dmsTone(tone: DmsTone): { bg: string; fg: string } {
  switch (tone) {
    case 'success': return { bg: clay.pillActiveBg, fg: clay.pillActiveText };
    case 'warning': return { bg: clay.pillWarnBg, fg: clay.pillWarnText };
    case 'danger': return { bg: clay.pillDangerBg, fg: clay.pillDangerText };
    case 'neutral': return { bg: 'var(--vdms-surface-alt)', fg: 'var(--vdms-text-muted)' };
    default: return { bg: clay.accentSoft, fg: clay.accent };
  }
}

/** Dashboard's exact page-title block (DashboardPage.tsx's `<h1>`/`<p>` pair)
 * — the master header spec every other page should render its own title
 * with, instead of a one-off heading. */
export function DmsPageHeader(props: { title: React.ReactNode; subtitle?: React.ReactNode; children?: React.ReactNode }): React.ReactElement {
  return (
    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: 12 }}>
      <div>
        <h1 style={{ margin: 0, fontSize: 'clamp(22px, 2.2vw, 28px)', fontWeight: 800, letterSpacing: '-0.02em', color: 'var(--vdms-text)', fontFamily: DMS_FONT_DISPLAY }}>
          {props.title}
        </h1>
        {props.subtitle && (
          <p style={{ margin: '3px 0 0', fontSize: 13, fontWeight: 600, color: 'var(--vdms-text-muted)' }}>{props.subtitle}</p>
        )}
      </div>
      {props.children && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>{props.children}</div>
      )}
    </div>
  );
}

/** Dashboard's `SectionCard` spec — the frosted, elevated container used for
 * every content block on Dashboard. Use for page-level card sections
 * elsewhere (not for the denser `DMS_TABLE_CARD` used by data tables). */
export const DMS_SECTION_CARD: React.CSSProperties = {
  background: 'var(--vdms-glass)', borderRadius: 18, border: '1px solid var(--vdms-line)',
  padding: 20, boxShadow: clay.shadowRaised,
  backdropFilter: 'blur(18px) saturate(1.3)', WebkitBackdropFilter: 'blur(18px) saturate(1.3)',
  display: 'flex', flexDirection: 'column', minWidth: 0,
};

/** Dashboard's `controlStyle()` — plain text input / native `<select>`. */
export function dmsControlStyle(): React.CSSProperties {
  return {
    padding: '8px 12px', borderRadius: 8, border: '1px solid var(--vdms-line)',
    fontSize: 12, background: 'var(--vdms-glass)', color: 'var(--vdms-text)',
    outline: 'none', minWidth: 0, boxShadow: clay.shadowRaised,
  };
}

/** Compact enterprise folder / library / vessel card (Folder view grids). */
export const DMS_TILE: React.CSSProperties = {
  background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-line)',
  padding: '12px 14px', display: 'flex', alignItems: 'center', gap: 12, cursor: 'pointer',
  boxShadow: clay.shadowRaised, minWidth: 0, boxSizing: 'border-box',
  transition: 'border-color 150ms ease, box-shadow 150ms ease, transform 150ms ease',
};
export const DMS_TILE_TITLE: React.CSSProperties = {
  fontWeight: 700, fontSize: 14, lineHeight: '20px', color: 'var(--vdms-text)',
  whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
};
export const DMS_TILE_SUB: React.CSSProperties = {
  fontSize: 12, fontWeight: 500, color: 'var(--vdms-text-muted)', marginTop: 2,
  display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap', minWidth: 0,
};
export const dmsGrid = (min: number): React.CSSProperties => ({
  display: 'grid', gridTemplateColumns: `repeat(auto-fill, minmax(${min}px, 1fr))`, gap: 12,
});

export function DmsTileIcon(props: { icon: string; tone?: DmsTone; size?: number }): React.ReactElement {
  const { bg, fg } = dmsTone(props.tone || 'accent');
  const size = props.size || 36;
  return (
    <div style={{ width: size, height: size, borderRadius: 10, background: bg, color: fg, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
      <Icon iconName={props.icon} style={{ fontSize: Math.round(size * 0.44) }} />
    </div>
  );
}

export function DmsChevron(): React.ReactElement {
  return <Icon iconName="ChevronRight" aria-hidden="true" style={{ fontSize: 11, color: 'var(--vdms-text-faint)', flexShrink: 0 }} />;
}

/** Folder / file count chip on a folder card. */
export function DmsCountChip(props: { icon: string; count: React.ReactNode; tone: DmsTone; on: boolean; label?: string }): React.ReactElement {
  const { bg, fg } = dmsTone(props.on ? props.tone : 'neutral');
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: bg, color: props.on ? fg : 'var(--vdms-text-faint)', borderRadius: 6, padding: '1px 7px', fontSize: 11, fontWeight: 700, lineHeight: '16px', whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums' }}>
      <Icon iconName={props.icon} aria-hidden="true" style={{ fontSize: 10 }} />
      {props.count}
      {props.label && <span style={{ fontWeight: 600, opacity: 0.8 }}>{props.label}</span>}
    </span>
  );
}

/** Section heading ("Folders", "Files", ...) with a soft count badge — mirrors Dashboard's section headers. */
export function DmsSectionLabel(props: { title: React.ReactNode; count?: number; tone?: DmsTone; children?: React.ReactNode }): React.ReactElement {
  const { bg, fg } = dmsTone(props.tone || 'accent');
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', minWidth: 0 }}>
      <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: 'var(--vdms-text)', fontFamily: DMS_FONT_DISPLAY, letterSpacing: '-0.01em' }}>{props.title}</h3>
      {typeof props.count === 'number' && (
        <span style={{ fontSize: 11, fontWeight: 700, color: fg, background: bg, borderRadius: 12, padding: '1px 8px', fontVariantNumeric: 'tabular-nums' }}>{props.count}</span>
      )}
      {props.children}
    </div>
  );
}

/** Clean empty / error state card (icon + message + optional existing action). */
export function DmsEmptyState(props: { icon: string; tone?: DmsTone; title: React.ReactNode; children?: React.ReactNode; action?: React.ReactNode }): React.ReactElement {
  return (
    <div style={{ background: 'var(--vdms-surface)', borderRadius: 16, border: '1px solid var(--vdms-line)', boxShadow: clay.shadowRaised, padding: '36px 20px', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6 }}>
      <DmsTileIcon icon={props.icon} tone={props.tone || 'accent'} size={44} />
      <div style={{ marginTop: 6, fontWeight: 700, color: 'var(--vdms-text)', fontSize: 15 }}>{props.title}</div>
      {props.children && <div style={{ fontSize: 13, color: 'var(--vdms-text-muted)', maxWidth: 460, lineHeight: 1.5 }}>{props.children}</div>}
      {props.action}
    </div>
  );
}

export function DmsSpinner(props: { size?: number }): React.ReactElement {
  const size = props.size || 28;
  return (
    <span aria-hidden="true" style={{ display: 'inline-block', width: size, height: size, borderRadius: '50%', border: `3px solid ${clay.accentSoft}`, borderTopColor: clay.accent, animation: 'dms-spin 0.8s linear infinite', boxSizing: 'border-box', flexShrink: 0 }} />
  );
}

/** Loading presentation only — callers still decide when it shows. */
export function DmsLoadingState(props: { label: React.ReactNode; hint?: React.ReactNode }): React.ReactElement {
  return (
    <div role="status" style={{ padding: '40px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10, textAlign: 'center' }}>
      <DmsSpinner />
      <div style={{ fontSize: 13.5, fontWeight: 600, color: 'var(--vdms-text)' }}>{props.label}</div>
      {props.hint && <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', maxWidth: 380 }}>{props.hint}</div>}
    </div>
  );
}

/** Toolbar button hierarchy: primary (solid accent), secondary (outlined), ghost (utility), danger. */
export type DmsBtnKind = 'primary' | 'secondary' | 'ghost' | 'danger';
export function dmsBtn(kind: DmsBtnKind, enabled = true): React.CSSProperties {
  const base: React.CSSProperties = {
    height: 34, padding: '0 14px', borderRadius: clay.radiusButton, boxSizing: 'border-box',
    fontSize: 13, fontWeight: 600, whiteSpace: 'nowrap', lineHeight: 1,
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7,
    cursor: enabled ? 'pointer' : 'not-allowed',
    transition: 'background 150ms ease, border-color 150ms ease, box-shadow 150ms ease, color 150ms ease',
  };
  if (!enabled) {
    return { ...base, background: kind === 'ghost' ? 'transparent' : 'var(--vdms-surface-alt)', color: 'var(--vdms-text-faint)', border: kind === 'ghost' ? '1px solid transparent' : '1px solid var(--vdms-border-soft)', boxShadow: 'none' };
  }
  switch (kind) {
    case 'primary': return { ...base, background: clay.accent, color: DMS_ON_ACCENT, border: '1px solid transparent', boxShadow: clay.shadowButton };
    case 'secondary': return { ...base, background: 'var(--vdms-surface)', color: 'var(--vdms-text)', border: '1px solid var(--vdms-line-strong)', boxShadow: '0 1px 2px rgba(16,27,45,0.04)' };
    case 'danger': return { ...base, background: clay.pillDangerBg, color: clay.pillDangerText, border: '1px solid transparent' };
    default: return { ...base, background: 'transparent', color: 'var(--vdms-text-secondary)', border: '1px solid transparent' };
  }
}

/** Folder-view file tables (library browser, Sites browser, vessel folder). */
export const DMS_TABLE_CARD: React.CSSProperties = {
  background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-line)',
  boxShadow: clay.shadowRaised, overflow: 'hidden',
};
export const DMS_TABLE: React.CSSProperties = { width: '100%', borderCollapse: 'collapse', fontSize: 13 };
export const DMS_TH: React.CSSProperties = {
  padding: '10px 14px', background: 'var(--vdms-surface-alt)', borderBottom: '1px solid var(--vdms-line)',
  color: 'var(--vdms-text-muted)', fontSize: 11, fontWeight: 700, letterSpacing: '0.05em',
  textTransform: 'uppercase', textAlign: 'left', whiteSpace: 'nowrap',
};
export const DMS_TR: React.CSSProperties = { borderBottom: '1px solid var(--vdms-border-soft)' };
export const DMS_TD: React.CSSProperties = { padding: '10px 14px', color: 'var(--vdms-text-muted)', fontSize: 12.5, verticalAlign: 'middle' };
export const DMS_TD_NAME: React.CSSProperties = { padding: '10px 14px', fontWeight: 600, color: 'var(--vdms-text)', verticalAlign: 'middle' };
export const DMS_NAME_CELL: React.CSSProperties = { display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 };
export const DMS_FILE_LINK: React.CSSProperties = { cursor: 'pointer', color: clay.accent, fontWeight: 600, textDecoration: 'none', wordBreak: 'break-word' };

/** Small in-row action buttons. */
export function dmsRowBtn(tone: DmsTone | 'plain'): React.CSSProperties {
  const base: React.CSSProperties = {
    height: 28, padding: '0 10px', borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: 'pointer',
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 5, whiteSpace: 'nowrap', boxSizing: 'border-box',
  };
  if (tone === 'plain') return { ...base, background: 'var(--vdms-surface)', color: 'var(--vdms-text)', border: '1px solid var(--vdms-line)' };
  const { bg, fg } = dmsTone(tone);
  return { ...base, background: bg, color: fg, border: '1px solid transparent' };
}

/**
 * "Classic Enterprise" compact dropdown — replaces a browser-native
 * `<select>`'s look with a custom popover (thin border, 6px radius, subtle
 * shadow, a 280px-max scrollable list of compact ~30px rows, accent-blue
 * selected row, soft hover) while keeping Fluent UI's `<Dropdown>` component
 * underneath, so `selectedKey`/`onChange`/`options` keep working exactly as
 * a normal controlled dropdown. Pass as `styles={styleProps => dmsCompactDropdownStyles(styleProps)}`.
 * First built for Documents' Vessel filter — reuse this for any other
 * single-select filter dropdown instead of a native `<select>`, so every
 * dropdown in the app opens the same compact, on-brand popover.
 */
export function dmsCompactDropdownStyles(styleProps: Partial<IDropdownStyleProps>, opts?: { width?: number }): Partial<IDropdownStyles> {
  const width = opts?.width ?? 220;
  return {
    root: { width, maxWidth: '100%', minWidth: 0, flex: `0 1 ${width}px` },
    dropdown: { width: '100%', minWidth: 0, fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 500 },    title: {
      height: 34, lineHeight: '32px', boxSizing: 'border-box', padding: '0 32px 0 10px',
      border: '1px solid var(--vdms-border)', borderRadius: 6,
      background: 'var(--vdms-surface)', color: 'var(--vdms-text)',
      fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 500, boxShadow: clay.shadowRaised, whiteSpace: 'nowrap',
    },
    caretDownWrapper: { height: 32, width: 28 },
    caretDown: {
      color: 'var(--vdms-text-muted)', fontSize: 12,
      transform: styleProps.isOpen ? 'rotate(180deg)' : 'none',
      transition: 'transform 120ms ease',
    },
    dropdownItemsWrapper: {
      maxHeight: 280, overflowY: 'auto', padding: '4px 0',
      scrollbarWidth: 'thin', scrollbarColor: 'var(--vdms-border) transparent',
      selectors: {
        '::-webkit-scrollbar': { width: 6 },
        '::-webkit-scrollbar-thumb': { background: 'var(--vdms-border)', borderRadius: 4 },
      },
    },
    dropdownItem: {
      height: 30, minHeight: 30, boxSizing: 'border-box', padding: '5px 12px',
      fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 500, lineHeight: '20px', color: 'var(--vdms-text)',
      selectors: {
        '&:hover': { background: 'var(--clay-surface-hover, #eaf0f6)', color: 'var(--vdms-text)' },
        '&:focus': { background: 'var(--clay-surface-hover, #eaf0f6)', color: 'var(--vdms-text)' },
      },
    },
    dropdownItemSelected: {
      height: 30, minHeight: 30, boxSizing: 'border-box', padding: '5px 12px',
      fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 500, lineHeight: '20px', background: clay.accent, color: '#ffffff',
      selectors: { '&:hover': { background: clay.accentHover, color: '#ffffff' } },
    },
    dropdownItemHeader: {
      height: 28, lineHeight: '28px', boxSizing: 'border-box', padding: '0 12px',
      fontSize: 11, fontWeight: 700, letterSpacing: '0.04em', textTransform: 'uppercase',
      color: 'var(--vdms-text-muted)',
    },
    dropdownOptionText: { fontFamily: vdmsFont.ui, fontSize: 15, fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
    callout: {
      maxHeight: 280, overflow: 'hidden', border: '1px solid var(--vdms-border)',
      borderRadius: 6, background: 'var(--vdms-surface)', boxShadow: clay.shadowRaised,
    },
  };
}

/** Pagination button (Folder view + List view footers). */
export function dmsPagerBtn(disabled: boolean, active = false): React.CSSProperties {
  return {
    minWidth: 28, height: 28, padding: '0 9px', borderRadius: 7, boxSizing: 'border-box',
    border: active ? '1px solid transparent' : '1px solid var(--vdms-line)',
    background: active ? clay.accent : 'var(--vdms-surface)',
    color: active ? DMS_ON_ACCENT : 'var(--vdms-text)',
    fontSize: 12, fontWeight: active ? 700 : 500,
    cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? 0.45 : 1,
  };
}
