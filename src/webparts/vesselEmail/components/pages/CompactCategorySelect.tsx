import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { dmsCompactDropdownStyles } from '../dmsDesignSystem';

interface CompactCategorySelectProps {
  'aria-label': string;
  value: string;
  onChange: (value: string) => void;
  style: React.CSSProperties;
  children: React.ReactNode;
  disabled?: boolean;
  title?: string;
}

function collectOptions(children: React.ReactNode, options: IDropdownOption[]): void {
  React.Children.forEach(children, child => {
    if (!React.isValidElement(child)) return;
    if (child.type === 'option') {
      options.push({ key: child.props.value, text: child.props.children, disabled: child.props.disabled });
    } else if (child.type === 'optgroup') {
      options.push({ key: `header:${child.key}`, text: child.props.label, itemType: DropdownMenuItemType.Header });
      collectOptions(child.props.children, options);
    } else if (child.type === React.Fragment) {
      collectOptions(child.props.children, options);
    }
  });
}

export function CompactCategorySelect(props: CompactCategorySelectProps): React.ReactElement {
  const options: IDropdownOption[] = [];
  collectOptions(props.children, options);
  return (
    <span title={props.title} style={{ position: 'relative', display: 'inline-flex', maxWidth: props.style.maxWidth }}>
      <select aria-hidden="true" tabIndex={-1} value={props.value} onChange={() => undefined}
        style={{ ...props.style, visibility: 'hidden', pointerEvents: 'none' }}>
        {props.children}
      </select>
      <Dropdown
        aria-label={props['aria-label']}
        selectedKey={props.value}
        options={options}
        onChange={(_, option) => { if (option) props.onChange(String(option.key)); }}
        dropdownWidth={0}
        disabled={props.disabled}
        calloutProps={props['aria-label'] === 'Category filter' ? { calloutMaxHeight: 300, className: 'dms-filter-callout' } : { className: 'dms-filter-callout' }}
        styles={styleProps => ({
          ...dmsCompactDropdownStyles(styleProps),
          ...(props['aria-label'] === 'Category filter' ? {
            dropdownItemsWrapper: { maxHeight: 240, overflowY: 'auto', padding: '2px 0' },
            dropdownItem: { minHeight: 28, height: 'auto', padding: '4px 10px', fontSize: 15, fontWeight: 500 },
            dropdownItemSelected: { minHeight: 28, height: 'auto', padding: '4px 10px', fontSize: 15, fontWeight: 500 },
            dropdownItemHeader: { padding: '4px 10px', fontSize: 13 },
          } : {}),
          // Overlay technique: the hidden native `<select>` above carries the
          // caller's real style (width/maxWidth/disabled opacity/etc.) and
          // reserves the layout box; this Dropdown absolutely fills it
          // (`inset: 0`) and does the actual rendering, so every existing
          // `<select>` in this file becomes compact/custom-styled just by
          // swapping its tag — no per-caller width/position bookkeeping.
          root: { position: 'absolute', inset: 0, minWidth: 0 },
          title: {
            height: 34, lineHeight: '32px', boxSizing: 'border-box',
            padding: '0 28px 0 10px', borderRadius: 6,
            border: '1px solid var(--vdms-border)', fontSize: 15, fontWeight: 500,
            background: props.disabled ? 'var(--vdms-surface-alt)' : 'var(--vdms-surface)',
            color: props.disabled ? 'var(--vdms-text-faint)' : 'var(--vdms-text)',
            cursor: props.disabled ? 'not-allowed' : 'pointer',
          },
          caretDownWrapper: { height: 32, width: 24 },
          caretDown: {
            fontSize: 10, color: 'var(--vdms-text-muted)',
            transform: styleProps.isOpen ? 'rotate(180deg)' : 'none',
            transition: 'transform 120ms ease',
          },
        })}
      />
    </span>
  );
}