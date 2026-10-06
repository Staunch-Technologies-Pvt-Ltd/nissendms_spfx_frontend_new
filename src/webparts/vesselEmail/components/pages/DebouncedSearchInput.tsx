import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';

/**
 * Debounced search text input shared by DocumentsPage.tsx (both the
 * dropdown-mode and panel-mode top-bar search boxes) and FilterPanel.tsx's
 * Keyword section.
 *
 * Why this exists: `textFilter` feeds straight into the Documents page's
 * expensive per-render recompute (the `groupCatActive` block — a recursive
 * Graph folder walk on cache-miss, and a synchronous BFS over up to 2000
 * cached folder entries on every render while that's loading). Previously
 * each raw <input onChange> called handleKeywordChange (host.setState) on
 * every keystroke, so that recompute — plus a full host re-render of the
 * ~7,000-line renderDocumentsPage function — ran once per character typed,
 * which is what made the search box feel like it "stopped working" while
 * typing quickly.
 *
 * This component keeps the box itself snappy (local state updates
 * immediately, so there's no input lag) but only calls the real onChange
 * (which flows into host.setState + host._scheduleGlobalSearch) 300ms after
 * the user stops typing.
 *
 * `value` is the source of truth for external resets — Clear All, leaving
 * and re-entering the Documents module (_goToView), switching document
 * sites/scopes, etc. all reset textFilter from outside this component, and
 * this component stays in sync with those.
 */
export interface DebouncedSearchInputProps {
  value: string;
  onChange: (val: string) => void;
  placeholder?: string;
  title?: string;
  style?: React.CSSProperties;
  className?: string;
  autoFocus?: boolean;
  /** Debounce delay in ms. Defaults to 300 (matches _scheduleGlobalSearch's 400ms closely enough that the backend lookup still lands just after the local recompute). */
  delayMs?: number;
  /** Forwarded to the underlying <input> — e.g. for a results dropdown that
   *  needs Escape/ArrowUp/ArrowDown/Enter, or open/close-on-focus/blur. */
  onKeyDown?: (e: React.KeyboardEvent<HTMLInputElement>) => void;
  onFocus?: (e: React.FocusEvent<HTMLInputElement>) => void;
  onBlur?: (e: React.FocusEvent<HTMLInputElement>) => void;
  /** Renders an inline "×" that clears the box immediately (bypassing the
   *  debounce) once there's text to clear. Defaults to true. */
  showClearButton?: boolean;
}

export function DebouncedSearchInput(props: DebouncedSearchInputProps): React.ReactElement {
  const {
    value, onChange, placeholder, title, style, className, autoFocus, delayMs,
    onKeyDown, onFocus, onBlur, showClearButton,
  } = props;
  const [local, setLocal] = React.useState(value);
  const timerRef = React.useRef<number | undefined>(undefined);
  // Tracks the last value we either received from outside or pushed out
  // ourselves, so an external reset (e.g. Clear All setting textFilter to
  // '') is picked up, but our own debounced onChange firing doesn't get
  // mistaken for an external change and bounce the local box.
  const lastKnownValue = React.useRef(value);

  React.useEffect(() => {
    if (value !== lastKnownValue.current) {
      lastKnownValue.current = value;
      setLocal(value);
    }
  }, [value]);

  React.useEffect(() => {
    return () => {
      if (timerRef.current !== undefined) {
        window.clearTimeout(timerRef.current);
      }
    };
  }, []);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const next = e.target.value;
    setLocal(next);
    if (timerRef.current !== undefined) {
      window.clearTimeout(timerRef.current);
    }
    timerRef.current = window.setTimeout(() => {
      lastKnownValue.current = next;
      onChange(next);
    }, delayMs ?? 300);
  };

  // Clears immediately rather than waiting out the debounce (Scenario:
  // "Clear search — Click X — Restore previous/default results").
  const handleClear = (): void => {
    if (timerRef.current !== undefined) {
      window.clearTimeout(timerRef.current);
      timerRef.current = undefined;
    }
    lastKnownValue.current = '';
    setLocal('');
    onChange('');
  };

  return (
    <span style={{ position: 'relative', display: 'block', width: '100%' }}>
      <input
        type="text"
        value={local}
        onChange={handleChange}
        onKeyDown={onKeyDown}
        onFocus={onFocus}
        onBlur={onBlur}
        placeholder={placeholder}
        title={title}
        style={style}
        className={className}
        autoFocus={autoFocus}
      />
      {showClearButton !== false && local && (
        <button
          type="button"
          // mousedown (not click) + preventDefault so clearing never steals
          // focus away from the input first.
          onMouseDown={e => e.preventDefault()}
          onClick={handleClear}
          aria-label="Clear search"
          title="Clear search"
          style={{
            position: 'absolute', right: 6, top: '50%', transform: 'translateY(-50%)',
            border: 'none', background: 'transparent', color: 'var(--vdms-text-faint)',
            fontSize: 13, lineHeight: 1, cursor: 'pointer', padding: 2,
          }}
        >
          <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 13 }} />
        </button>
      )}
    </span>
  );
}
