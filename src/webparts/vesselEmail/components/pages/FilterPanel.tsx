import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DebouncedSearchInput } from './DebouncedSearchInput';

/**
 * Documents page — slide-out "Filters" panel.
 *
 * Alternative presentation for the existing Documents filters (Main
 * folder, Vessel, Category, Sub-folder, Attachment status), grouped into
 * collapsible sections with a keyword box up top and Clear All / Apply
 * buttons at the bottom — the layout referenced in Settings → Filter
 * Search Management ("panel" mode, filter_settings_api.py).
 *
 * This component is presentation-only: every value/onChange it's given
 * comes straight from DocumentsPage.tsx's existing filter state and the
 * exact same handlers the inline dropdown row uses (docGroupFilter,
 * vesselFilter, docCategoryFilter, docLeafCategoryFilter,
 * attachmentFilter, docSubfolderOtherFilter, ...) — so switching between
 * "dropdown" and "panel" in Settings never changes what a filter does,
 * only how it's presented. Every option list here is single-select
 * (radio), matching the underlying state, which only ever holds one value
 * per filter — not a multi-select checkbox list.
 */

export interface FilterPanelOption { value: string; label: string; }

export interface FilterPanelSection {
  key: string;
  label: string;
  value: string;
  options: FilterPanelOption[];
  onChange: (val: string) => void;
  /** Optional hint shown under the section label (e.g. why it's scoped). */
  hint?: string;
}

export interface FilterPanelProps {
  open: boolean;
  onClose: () => void;
  onClearAll: () => void;
  keyword: { value: string; onChange: (val: string) => void };
  sections: FilterPanelSection[];
  /** Rendered as its own "Sub-folder" section — the live folder tree
   *  (FolderTreeSelect) doesn't fit the flat radio-list shape the other
   *  sections use, so the caller renders it and hands it in as-is. */
  subfolderSection?: React.ReactNode;
  activeCount: number;
}

const SECTION_LABEL: React.CSSProperties = {
  fontSize: 12.5, fontWeight: 700, color: 'var(--vdms-text)', display: 'flex',
  alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer',
  padding: '10px 2px', userSelect: 'none',
};

function Section({ title, hint, children, defaultOpen }: { title: string; hint?: string; children: React.ReactNode; defaultOpen?: boolean }): React.ReactElement {
  return (
    <details open={defaultOpen !== false} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
      <summary style={SECTION_LABEL}>
        <span>
          {title}
          {hint && <span style={{ display: 'block', fontSize: 11, fontWeight: 500, color: 'var(--vdms-text-muted)', marginTop: 2 }}>{hint}</span>}
        </span>
      </summary>
      <div style={{ paddingBottom: 12, display: 'flex', flexDirection: 'column', gap: 6 }}>
        {children}
      </div>
    </details>
  );
}

function RadioOption({ name, value, label, checked, onSelect }: { name: string; value: string; label: string; checked: boolean; onSelect: () => void }): React.ReactElement {
  return (
    <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12.5, color: 'var(--vdms-text)', cursor: 'pointer', padding: '2px 2px' }}>
      <input type="radio" name={name} value={value} checked={checked} onChange={onSelect} />
      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</span>
    </label>
  );
}

export function FilterPanel(props: FilterPanelProps): React.ReactElement | null {
  const { open, onClose, onClearAll, keyword, sections, subfolderSection, activeCount } = props;
  if (!open) return null;

  return (
    <>
      {/* Overlay — click to dismiss, like the reference Filters panel. */}
      <div
        onClick={onClose}
        style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.35)', zIndex: 1000 }}
      />
      <div
        role="dialog"
        aria-label="Filters"
        style={{
          position: 'fixed', top: 0, right: 0, bottom: 0, width: 320, maxWidth: '92vw',
          background: 'var(--vdms-surface)', boxShadow: '-6px 0 24px rgba(15,23,42,0.18)',
          zIndex: 1001, display: 'flex', flexDirection: 'column',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 16px', borderBottom: '1px solid var(--vdms-border)' }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: 'var(--vdms-text)' }}>Filters</div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close filters"
            style={{ border: 'none', background: 'transparent', fontSize: 16, cursor: 'pointer', color: 'var(--vdms-text-muted)', lineHeight: 1, padding: 4 }}
          >
            <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 16 }} />
          </button>
        </div>

        <div style={{ flex: '1 1 auto', overflowY: 'auto', padding: '4px 16px' }}>
          <Section title="Keyword">
            <DebouncedSearchInput
              placeholder="Search vessel, file name, folder, group, category..."
              value={keyword.value}
              onChange={keyword.onChange}
              style={{ width: '100%', padding: '7px 26px 7px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12.5, outline: 'none', boxSizing: 'border-box' }}
            />
          </Section>

          {sections.map(section => (
            <Section key={section.key} title={section.label} hint={section.hint}>
              {section.options.map(opt => (
                <RadioOption
                  key={opt.value || '(empty)'}
                  name={`filter-panel-${section.key}`}
                  value={opt.value}
                  label={opt.label}
                  checked={section.value === opt.value}
                  onSelect={() => section.onChange(opt.value)}
                />
              ))}
            </Section>
          ))}

          {subfolderSection && (
            <Section title="Sub-folder">
              {subfolderSection}
            </Section>
          )}
        </div>

        <div style={{ display: 'flex', gap: 8, padding: '12px 16px', borderTop: '1px solid var(--vdms-border)' }}>
          <button
            type="button"
            onClick={onClearAll}
            disabled={activeCount === 0}
            style={{
              flex: '1 1 0', padding: '8px 0', borderRadius: 8, border: '1px solid var(--vdms-border)',
              background: 'var(--vdms-surface)', color: activeCount === 0 ? 'var(--vdms-text-faint)' : 'var(--vdms-text)',
              fontSize: 12.5, fontWeight: 600, cursor: activeCount === 0 ? 'not-allowed' : 'pointer',
            }}
          >
            Clear All
          </button>
          <button
            type="button"
            onClick={onClose}
            style={{
              flex: '1 1 0', padding: '8px 0', borderRadius: 8, border: 'none',
              background: '#0a66d0', color: '#fff', fontSize: 12.5, fontWeight: 700, cursor: 'pointer',
            }}
          >
            Apply
          </button>
        </div>
      </div>
    </>
  );
}
