import * as React from 'react';

/** One node of a vessel's live sub-folder hierarchy. `id` is the live
 *  SharePoint folder id when known, otherwise a synthetic key (the folder's
 *  own name) for a folder that was only discovered via row-scanning and has
 *  no confirmed nesting — it's rendered as a childless top-level entry, same
 *  as the old flat dropdown did for that case. */
export interface FolderTreeNode {
  name: string;
  id: string;
  children: FolderTreeNode[];
}

interface FolderTreeSelectProps {
  /** Root nodes of the tree, already scoped to the selected vessel. */
  tree: FolderTreeNode[];
  /** Currently selected folder name, or 'all'. Matches the plain-<select>
   *  contract this replaces — docSubfolderOtherFilter's value shape. */
  value: string;
  /** Exact path (names from the root) of the selected node, when known. Lets
   *  two folders sharing a name (Drawings > Electrical vs Manuals > Electrical)
   *  be told apart — without it, selection falls back to matching by name. */
  selectedPath?: string[];
  /** `name` is the selected node's own name (matches the plain-<select>
   *  contract — this is what gets stored as the filter value). `path` is the
   *  full chain of names from the tree's roots down to and including this
   *  node, for callers that need to navigate to it (a name alone doesn't say
   *  where in the tree it lives when the same name could appear at more than
   *  one level). */
  onChange: (name: string, path: string[], id?: string) => void;
  /** Optional lazy loader. When given, a folder whose children aren't in
   *  `tree` yet still shows the expand arrow (if `canLoadChildren` allows);
   *  opening it fetches that folder's direct sub-folders on demand, so every
   *  nested level is reachable from the arrow without navigating into each
   *  folder one by one — and without walking the whole tree up front. */
  loadChildren?: (node: FolderTreeNode, path: string[]) => Promise<FolderTreeNode[]>;
  canLoadChildren?: (node: FolderTreeNode) => boolean;
  placeholder?: string;
  allLabel?: string;
  title?: string;
  disabled?: boolean;
  style?: React.CSSProperties;
}

/** Depth-first walk collecting every ancestor path (as an array of node
 *  keys) leading to a node whose name matches `query`. Used to auto-expand
 *  the tree down to search hits without hand-rolling recursion at each call
 *  site. */
function collectMatchPaths(nodes: FolderTreeNode[], query: string, trail: string[], out: Set<string>): boolean {
  let anyMatch = false;
  for (const node of nodes) {
    const key = [...trail, node.id].join('/');
    const selfMatch = node.name.toLowerCase().includes(query);
    const childMatch = collectMatchPaths(node.children, query, [...trail, node.id], out);
    if (selfMatch || childMatch) {
      anyMatch = true;
      // Every ancestor on the way down to a hit needs to be expanded — add
      // this node's own key so it renders open when it's someone's parent.
      out.add(key);
    }
  }
  return anyMatch;
}

function nodeOrDescendantMatches(node: FolderTreeNode, query: string): boolean {
  if (node.name.toLowerCase().includes(query)) return true;
  return node.children.some(c => nodeOrDescendantMatches(c, query));
}

/** Single-select dropdown for a folder hierarchy, styled to match the plain
 *  filter `<select>`s next to it but backed by a popup tree panel instead of
 *  native `<option>`s — a native select can't expand/collapse or nest rows,
 *  which is what a multi-level folder tree (Thrusters → Engine → NK →
 *  Stamp/Poutfit/...) needs. Selecting any node, heading or leaf, sets the
 *  filter to that node's name; the existing row-filtering logic already
 *  matches a path SEGMENT anywhere in a row's folder path, so picking a
 *  heading already includes everything nested under it. */
export function FolderTreeSelect(props: FolderTreeSelectProps): React.ReactElement {
  const { tree, value, onChange, disabled, loadChildren, canLoadChildren } = props;
  const placeholder = props.allLabel || props.placeholder || 'All sub-folders';
  const [open, setOpen] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [manualExpanded, setManualExpanded] = React.useState<Set<string>>(new Set());
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const searchRef = React.useRef<HTMLInputElement | null>(null);

  React.useEffect(() => {
    if (!open) return;
    const onDocMouseDown = (e: MouseEvent): void => {
      if (rootRef.current && e.target instanceof Node && !rootRef.current.contains(e.target)) {
        setOpen(false);
      }
    };
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDocMouseDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onDocMouseDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open]);

  React.useEffect(() => {
    if (open) {
      // Popup opens collapsed at top level and with an empty search, every
      // time — matches the "default state: collapsed at top level" spec
      // rather than remembering the last session's expand state.
      setSearch('');
      const t = setTimeout(() => searchRef.current?.focus(), 0);
      return () => clearTimeout(t);
    }
    return undefined;
  }, [open]);

  // Lazily fetched children, keyed like `manualExpanded` (ancestor id trail).
  // `null` = a load finished and found no sub-folders (arrow goes away).
  const [lazyChildren, setLazyChildren] = React.useState<Record<string, FolderTreeNode[] | null>>({});
  const [loadingKeys, setLoadingKeys] = React.useState<Set<string>>(new Set());
  const inFlight = React.useRef<Set<string>>(new Set());

  const normSearch = search.trim().toLowerCase();
  const searchExpanded = React.useMemo(() => {
    if (!normSearch) return new Set<string>();
    const out = new Set<string>();
    collectMatchPaths(tree, normSearch, [], out);
    return out;
  }, [tree, normSearch]);

  // Tree is always fully expanded — no arrows/dropdowns to click.
  const isExpanded = (_key: string): boolean => true;
  const mergeChildren = (own: FolderTreeNode[], extra: FolderTreeNode[] | null | undefined): FolderTreeNode[] => {
    if (!extra || extra.length === 0) return own;
    const seen = new Set(own.map(c => c.name.trim().toLowerCase()));
    return [...own, ...extra.filter(c => !seen.has(c.name.trim().toLowerCase()))];
  };
  const lazyEligible = (node: FolderTreeNode): boolean =>
    !!loadChildren && (!canLoadChildren || canLoadChildren(node));
  const ensureLoaded = (node: FolderTreeNode, key: string, path: string[]): void => {
    if (!lazyEligible(node) || key in lazyChildren || inFlight.current.has(key)) return;
    inFlight.current.add(key);
    setLoadingKeys(prev => new Set(prev).add(key));
    loadChildren!(node, path)
      .then(kids => setLazyChildren(prev => ({ ...prev, [key]: kids && kids.length > 0 ? kids : null })))
      .catch(() => setLazyChildren(prev => ({ ...prev, [key]: null })))
      .then(() => {
        inFlight.current.delete(key);
        setLoadingKeys(prev => { const n = new Set(prev); n.delete(key); return n; });
      });
  };
  // While the popup is open, eagerly load every nested level (a few requests
  // at a time) so the whole hierarchy becomes visible without any clicking.
  React.useEffect(() => {
    if (!open || !loadChildren) return;
    const MAX_IN_FLIGHT = 6;
    const walk = (nodes: FolderTreeNode[], trail: string[], namePath: string[]): void => {
      for (const node of nodes) {
        if (inFlight.current.size >= MAX_IN_FLIGHT) return;
        const key = [...trail, node.id].join('/');
        const ownPath = [...namePath, node.name];
        ensureLoaded(node, key, ownPath);
        walk(mergeChildren(node.children, lazyChildren[key]), [...trail, node.id], ownPath);
      }
    };
    walk(tree, [], []);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, tree, lazyChildren]);

  const select = (name: string, path: string[], id?: string): void => {
    onChange(name, path, id);
    setOpen(false);
  };

  const countDescendants = (node: FolderTreeNode): number =>
    node.children.reduce((sum, c) => sum + 1 + countDescendants(c), 0);

  const renderNodes = (nodes: FolderTreeNode[], trail: string[], namePath: string[], depth: number): React.ReactElement[] => {
    const visible = normSearch ? nodes.filter(n => nodeOrDescendantMatches(n, normSearch)) : nodes;
    const railLeft = 8 + depth * 16 + 9;
    return visible.map(node => {
      const key = [...trail, node.id].join('/');
      const ownPath = [...namePath, node.name];
      const expanded = isExpanded(key);
      // Merge in anything fetched on demand for this node.
      const lazyKids = lazyChildren[key];
      const kids = mergeChildren(node.children, lazyKids);
      const loading = loadingKeys.has(key);
      // Arrow shows for known children, or (not yet loaded) any live folder
      // that could have some — it disappears again if the load finds none.
      const hasChildren = kids.length > 0 || loading || (lazyEligible(node) && !(key in lazyChildren));
      const isSelected = value !== 'all' && value.trim().toLowerCase() === node.name.trim().toLowerCase()
        && (!props.selectedPath || props.selectedPath.map(n => n.trim().toLowerCase()).join('/') === ownPath.map(n => n.trim().toLowerCase()).join('/'));
      const selfMatches = !normSearch || node.name.toLowerCase().includes(normSearch);
      return (
        <div key={key} style={{ position: 'relative' }}>
          {/* Vertical guide connecting this row to its parent's toggle column,
             so a deeply nested folder still reads as belonging to its branch
             instead of floating at an arbitrary indent. */}
          {depth > 0 && (
            <div
              aria-hidden="true"
              style={{ position: 'absolute', left: railLeft - 16, top: 0, bottom: 0, width: 1, background: 'var(--vdms-border)', opacity: 0.6 }}
            />
          )}
          <div
            className="fts-row"
            style={{
              display: 'flex', alignItems: 'center', gap: 6, padding: '6px 8px',
              paddingLeft: 8 + depth * 16, borderRadius: 6, cursor: 'pointer',
              ...(isSelected ? { background: '#fff7ed' } : {}),
              position: 'relative',
            }}
            onMouseDown={e => e.preventDefault()}
            onClick={() => select(node.name, ownPath, node.id)}
            title={node.name}
          >
            {/* Branch connector instead of an expand arrow — nesting is always shown. */}
            <span
              aria-hidden="true"
              style={{ width: 14, flexShrink: 0, textAlign: 'center', fontSize: 11, color: 'var(--vdms-text-muted)' }}
            >
              {depth > 0 ? '└' : ''}
            </span>
            <span style={{ fontSize: 13, flexShrink: 0, opacity: hasChildren ? 1 : 0.7 }}>
              {kids.length > 0 ? '📂' : '📁'}
            </span>
            <span
              style={{
                fontSize: 12.5, flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                color: isSelected ? '#c2410c' : (depth === 0 ? '#111827' : 'var(--vdms-text)'),
                fontWeight: isSelected ? 700 : (depth === 0 ? 700 : 400),
                opacity: normSearch && !selfMatches ? 0.55 : 1,
              }}
              title={node.name}
            >
              {node.name}
            </span>
            {loading && (
              <span style={{ fontSize: 11, color: 'var(--vdms-text-muted)', flexShrink: 0 }}>…</span>
            )}
            {kids.length > 0 && (
              <span
                style={{
                  fontSize: 10.5, fontWeight: 600, color: 'var(--vdms-text-muted)', background: 'var(--vdms-surface-alt)',
                  borderRadius: 999, padding: '1px 6px', flexShrink: 0,
                }}
              >
                {countDescendants({ ...node, children: kids })}
              </span>
            )}
          </div>
          {hasChildren && expanded && (
            <div>{renderNodes(kids, [...trail, node.id], ownPath, depth + 1)}</div>
          )}
        </div>
      );
    });
  };

  const selectedLabel = value !== 'all' ? value : placeholder;

  return (
    <div ref={rootRef} style={{ position: 'relative', display: 'inline-block' }}>
      <button
        type="button"
        disabled={disabled}
        onClick={() => setOpen(o => !o)}
        title={props.title}
        style={{
          padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12,
          background: disabled ? 'var(--vdms-surface-alt)' : 'var(--vdms-surface)', outline: 'none',
          maxWidth: 180, color: disabled ? 'var(--vdms-text-faint)' : 'var(--vdms-text)',
          cursor: disabled ? 'not-allowed' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6,
          ...(props.style || {}),
        }}
      >
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1, textAlign: 'left' }}>
          {selectedLabel}
        </span>
        <span style={{ fontSize: 10, flexShrink: 0 }}>▾</span>
      </button>

      {open && !disabled && (
        <div
          style={{
            position: 'absolute', top: 'calc(100% + 4px)', left: 0, zIndex: 1000,
            width: 360, maxHeight: 480, display: 'flex', flexDirection: 'column',
            background: 'var(--vdms-surface)', border: '1px solid var(--vdms-border)', borderRadius: 10,
            boxShadow: '0 8px 28px rgba(0,0,0,0.18)', overflow: 'hidden',
          }}
        >
          <style>{`
            .fts-row:hover { background: var(--vdms-surface-alt); }
            .fts-toggle:hover { background: var(--vdms-border); }
          `}</style>
          <div style={{ padding: 8, borderBottom: '1px solid var(--vdms-border)', flexShrink: 0 }}>
            <input
              ref={searchRef}
              type="text"
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="Search folders…"
              style={{
                width: '100%', boxSizing: 'border-box', borderRadius: 6, border: '1px solid var(--vdms-border)',
                padding: '6px 8px', fontSize: 12.5, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)', outline: 'none',
              }}
            />
          </div>
          <div style={{ overflowY: 'auto', padding: 4 }}>
            <div
              style={{
                display: 'flex', alignItems: 'center', padding: '5px 8px', borderRadius: 6, cursor: 'pointer',
                background: value === 'all' ? '#fff7ed' : 'transparent',
                fontSize: 12.5, fontWeight: value === 'all' ? 700 : 600, color: value === 'all' ? '#c2410c' : 'var(--vdms-text)',
              }}
              onMouseDown={e => e.preventDefault()}
              onClick={() => select('all', [])}
            >
              {placeholder}
            </div>
            {tree.length === 0 ? (
              <div style={{ padding: '10px 8px', fontSize: 12, color: 'var(--vdms-text-faint)' }}>No sub-folders found.</div>
            ) : normSearch && !tree.some(n => nodeOrDescendantMatches(n, normSearch)) ? (
              <div style={{ padding: '10px 8px', fontSize: 12, color: 'var(--vdms-text-faint)' }}>No folders match "{search}".</div>
            ) : (
              renderNodes(tree, [], [], 0)
            )}
          </div>
        </div>
      )}
    </div>
  );
}
