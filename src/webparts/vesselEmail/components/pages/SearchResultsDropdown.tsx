import * as React from 'react';

/**
 * Live results list rendered directly under the Documents page's search
 * box ("provide the results in the search bar" — as opposed to only
 * filtering the file/folder list further down the page). Purely
 * presentational: DocumentsPage.tsx builds `items` from the same
 * matchesSearchTokens matcher the list view already filters with, so a
 * suggestion here is never inconsistent with what actually shows once the
 * debounce fires.
 */

export interface SearchDropdownItem {
  key: string;
  kind: 'vessel' | 'folder' | 'file';
  label: string;
  sublabel?: string;
  /** Called on click/Enter — navigates to the vessel/folder or opens the file. */
  onSelect: () => void;
}

export interface SearchResultsDropdownProps {
  open: boolean;
  query: string;
  items: SearchDropdownItem[];
  /** Index of the keyboard-highlighted row, or -1 for none. */
  activeIndex: number;
  highlight?: (text: string | null | undefined, query: string) => React.ReactNode;
}

const KIND_ICON: Record<SearchDropdownItem['kind'], string> = {
  vessel: '🚢', folder: '📁', file: '📄',
};

export function SearchResultsDropdown(props: SearchResultsDropdownProps): React.ReactElement | null {
  const { open, query, items, activeIndex, highlight } = props;
  const trimmed = (query || '').trim();
  if (!open || !trimmed) return null;

  return (
    <div
      role="listbox"
      aria-label="Search results"
      style={{
        position: 'absolute', top: '100%', left: 0, right: 0, marginTop: 4, zIndex: 40,
        background: 'var(--vdms-surface)', border: '1px solid var(--vdms-border)', borderRadius: 8,
        boxShadow: '0 8px 24px rgba(0,0,0,0.14)', maxHeight: 320, overflowY: 'auto', fontSize: 12.5,
      }}
    >
      {items.length === 0 ? (
        <div style={{ padding: '12px 14px', color: 'var(--vdms-text-faint)' }}>
          No results found for &quot;{trimmed}&quot;
        </div>
      ) : (
        items.map((item, idx) => (
          <div
            key={item.key}
            role="option"
            aria-selected={idx === activeIndex}
            // Prevent-default on mousedown (not click) so the input never
            // blurs before onSelect runs — a plain onClick would lose the
            // selection the moment the box's onBlur closes this dropdown.
            onMouseDown={e => { e.preventDefault(); item.onSelect(); }}
            style={{
              display: 'flex', alignItems: 'center', gap: 8, padding: '8px 14px', cursor: 'pointer',
              background: idx === activeIndex ? 'var(--vdms-border-soft, #f1f5f9)' : 'transparent',
              borderBottom: idx === items.length - 1 ? 'none' : '1px solid var(--vdms-border-soft, #f1f5f9)',
            }}
          >
            <span aria-hidden="true">{KIND_ICON[item.kind]}</span>
            <span style={{ display: 'flex', flexDirection: 'column', minWidth: 0, flex: 1 }}>
              <span style={{
                fontWeight: 600, color: 'var(--vdms-text)', overflow: 'hidden',
                textOverflow: 'ellipsis', whiteSpace: 'nowrap',
              }}>
                {highlight ? highlight(item.label, query) : item.label}
              </span>
              {item.sublabel && (
                <span style={{ fontSize: 11, color: 'var(--vdms-text-faint)' }}>{item.sublabel}</span>
              )}
            </span>
          </div>
        ))
      )}
    </div>
  );
}
