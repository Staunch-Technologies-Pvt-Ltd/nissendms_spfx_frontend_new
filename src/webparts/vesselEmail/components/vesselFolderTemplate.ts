/**
 * vesselFolderTemplate.ts
 *
 * Single source-of-truth for the Vessel Management folder structure.
 *
 * Shape:
 *   FolderNode.name        — folder display name
 *   FolderNode.children    — nested sub-folders (omit or [] for leaf)
 *   FolderNode.perVessel   — true  → created once per vessel under {VesselName}/
 *                            false → created once globally (common folders)
 *
 * ⚠️  Folder-4 (Kaizen – Knowledge Bank) is intentionally EXCLUDED until the
 *     business owner confirms the exact subfolder structure from the source PDF.
 *     Add it here once confirmed; no other file needs to change.
 */

export interface FolderNode {
  name: string;
  children?: FolderNode[];
}

export interface MainFolder {
  /** SharePoint folder name, e.g. "Folder-1 Technical & Crewing" */
  name: string;
  /** Sub-tree created once per vessel under {VesselName}/ */
  perVesselTree: FolderNode[];
  /** Sub-tree created once globally (common folders) */
  commonTree: FolderNode[];
  /**
   * Override the SharePoint folder name used for the "common" folder inside
   * this department. Defaults to COMMON_SHIPS_ROOT ('Common for all ships')
   * when not specified.
   */
  commonFolderName?: string;
}

// ── Folder-1: Technical & Crewing ────────────────────────────────────────────

const TECHNICAL_CREWING_PER_VESSEL: FolderNode[] = [
  {
    name: 'Month End Reports',
    children: [
      { name: 'Main Engine' },
      { name: 'Aux Engine' },
      { name: 'Cooling Water' },
      { name: 'Inspection Reports' },
      { name: 'Defect Reports' },
      { name: 'Guarantee Claims' },
      { name: 'To be Classified' },
    ],
  },
  {
    name: 'Service Agreements',
    children: [
      { name: 'Technical Management' },
      { name: 'Crew Management' },
      { name: 'Vendor & Service Provider' },
      { name: 'To be Classified' },
    ],
  },
  {
    name: 'Registration',
    children: [
      { name: 'Flag & MPA' },
      { name: 'Ship Builder' },
      { name: 'Radio & Telecom' },
      { name: 'Crewing & SMOU' },
      { name: 'Novation' },
      { name: 'To be Classified' },
    ],
  },
  {
    name: 'Drawings and Manuals',
    children: [
      {
        name: 'Drawing',
        children: [
          { name: 'Archive' },
          { name: 'Basic' },
          { name: 'Electrical' },
          { name: 'Hull' },
          { name: 'Machinery' },
          { name: 'Other Drawings' },
          { name: 'Safety' },
        ],
      },
      {
        name: 'Manual',
        children: [
          { name: 'Automation' },
          { name: 'Auxiliary Engine' },
          { name: 'Boiler' },
          { name: 'Cargo' },
          { name: 'Deck Machinery' },
          { name: 'Electrical' },
          { name: 'Main Engine' },
          { name: 'Other Manuals' },
          { name: 'Pollution' },
          { name: 'Propulsion' },
          { name: 'Refrigeration' },
          { name: 'Safety' },
          { name: 'Shafting' },
          { name: 'Steering Gear' },
          { name: 'Thrusters' },
        ],
      },
      { name: 'To be Classified' },
    ],
  },
   {
    name: 'PO & Invoice',
    children: [
      { name: 'Purchase Order' },
      { name: 'Vendor Invoice' },
    ],
  },
  { name: 'Incidents' },
  { name: 'Crewing' },
  { name: 'To be Classified' },
];

const TECHNICAL_CREWING_COMMON: FolderNode[] = [
  {
    name: 'Vendor & Service Agreements',
    children: [
      { name: 'Vendor & Service Provider Agreement' },
      { name: 'To be Classified' },
    ],
  },
  { name: 'Vendor Management' },
  { name: 'To be Classified' },
];

// ── Folder-2: Commercial and Chartering ──────────────────────────────────────

const COMMERCIAL_CHARTERING_PER_VESSEL: FolderNode[] = [
  {
    name: 'Agreements',
    children: [
      { name: 'Charter party' },
      { name: 'Pool Agreement' },
      { name: 'Commission Agreement' },
      { name: 'To be Classified' },
    ],
  },
  {
    name: 'Invoices & Payments',
    children: [
      { name: 'Invoice' },
      { name: 'Payments' },
      { name: 'To be Classified' },
    ],
  },
  {
    name: 'Claims & Disputes',
    children: [
      { name: 'Disputes' },
      { name: 'Claims' },
      { name: 'To be Classified' },
    ],
  },
  { name: 'To be Classified' },
];

const COMMERCIAL_CHARTERING_COMMON: FolderNode[] = [
  {
    name: 'Agreements',
    children: [
      { name: 'Charter party' },
      { name: 'Pool Agreement' },
      { name: 'Commission Agreement' },
      { name: 'To be Classified' },
    ],
  },
  { name: 'To be Classified' },
];

// ── Folder-3: Insurance ───────────────────────────────────────────────────────

const INSURANCE_PER_VESSEL: FolderNode[] = [
  { name: 'P&I' },
  { name: 'H&M' },
  { name: 'War Risk' },
  { name: 'Flag & MPA' },
  { name: 'USA Related' },
];

const INSURANCE_COMMON: FolderNode[] = [
  { name: 'Agreements' },
  { name: 'Miscellaneous' },
];

// ── Kaizen - Knowledge Bank (Global / Standalone at Documents Root) ───────────

export const KAIZEN_KNOWLEDGE_BANK_TREE: FolderNode[] = [
  { name: 'Templates' },
  {
    name: 'Procedures and Work Instructions',
    children: [
      { name: 'Equipment Maker' },
      { name: 'Class' },
      { name: 'Flag - Port State' },
      { name: 'SIRE-OCIMF-RightShip' },
      { name: 'Shipyard' },
    ],
  },
  { name: 'Lessons Learned' },
  {
    name: 'Circulars and Guidance',
    children: [
      { name: 'Equipment Maker' },
      { name: 'Class' },
      { name: 'Flag - Port State' },
      { name: 'SIRE-OCIMF-RightShip' },
      { name: 'Shipyard' },
    ],
  },
];

export const KAIZEN_KNOWLEDGE_BANK_COMMON = KAIZEN_KNOWLEDGE_BANK_TREE;

// ── Exported template ─────────────────────────────────────────────────────────

// Root constants (Main folders and Kaizen are directly at Documents root)
export const VESSEL_MANAGEMENT_ROOT = '';
export const SPECIFIC_VESSELS_ROOT = '';
export const COMMON_SHIPS_ROOT = 'Common for all ships';

export const MAIN_FOLDERS: MainFolder[] = [
  {
    name: 'Technical & Crewing',
    perVesselTree: TECHNICAL_CREWING_PER_VESSEL,
    commonTree: TECHNICAL_CREWING_COMMON,
  },
  {
    name: 'Commercial & Chartering',
    perVesselTree: COMMERCIAL_CHARTERING_PER_VESSEL,
    commonTree: COMMERCIAL_CHARTERING_COMMON,
    commonFolderName: 'Common Agreements (Not Ship Specific)',
  },
  {
    name: 'Insurance',
    perVesselTree: INSURANCE_PER_VESSEL,
    commonTree: INSURANCE_COMMON,
    commonFolderName: 'Common (Not Ship Specific)',
  },
];

// NOTE: an earlier pass in this session deleted the functions below as
// "dead code" based on a grep that missed DocumentsPage.tsx / FolderView.tsx
// / ListView.tsx — those three files existed on the device but had never
// been staged into that session, so the grep couldn't see their imports.
// They are NOT dead: DocumentsPage.tsx (the actual Document Library /
// flattened-tree view) and FolderView.tsx (the Folder View's Level-0/1
// navigation) both call every one of these. Restored as-is.

function dedupeNames(names: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  names.forEach(name => {
    const key = (name || '').trim().toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(name);
  });
  return out;
}

/** Display-ready folder lookups used by Documents and Folder view. */
export function folderNamesByMainFolder(common: boolean = false): Record<string, string[]> {
  const map = Object.fromEntries(MAIN_FOLDERS.map(main => [
    main.name,
    dedupeNames((common ? main.commonTree : main.perVesselTree).map(folder => folder.name)),
  ]));
  map['Kaizen - Knowledge Bank'] = dedupeNames(KAIZEN_KNOWLEDGE_BANK_TREE.map(f => f.name));
  return map;
}

export function subfolderNamesByFolder(): Record<string, string[]> {
  const result: Record<string, string[]> = {};
  const visit = (folders: FolderNode[], pathPrefix: string = ''): void => {
    for (const folder of folders) {
      const fullPath = pathPrefix ? `${pathPrefix} > ${folder.name}` : folder.name;
      if (folder.children?.length) {
        const childNames = dedupeNames(folder.children.map(child => child.name));
        result[fullPath] = childNames;
        // Only set the un-prefixed folder.name if not already set by a prior branch
        if (!result[folder.name]) {
          result[folder.name] = childNames;
        }
        visit(folder.children, fullPath);
      }
    }
  };
  for (const main of MAIN_FOLDERS) {
    visit(main.perVesselTree, main.name);
    visit(main.commonTree, main.name);
  }
  visit(KAIZEN_KNOWLEDGE_BANK_TREE, 'Kaizen - Knowledge Bank');
  return result;
}

/** Returns all folder names at depth >= 2 (i.e. child subfolders and nested leaves, not top categories). */
export function getAllKnownNestedTemplateSubfolders(mainFolderName?: string): Set<string> {
  const nested = new Set<string>();
  const collectChildren = (nodes: FolderNode[], isTopLevel: boolean): void => {
    for (const node of nodes) {
      if (!isTopLevel) {
        nested.add(node.name.trim().toLowerCase());
      }
      if (node.children?.length) {
        collectChildren(node.children, false);
      }
    }
  };

  const mainsToScan = mainFolderName
    ? MAIN_FOLDERS.filter(m => m.name.toLowerCase() === mainFolderName.toLowerCase())
    : MAIN_FOLDERS;

  for (const main of mainsToScan) {
    collectChildren(main.perVesselTree, true);
    collectChildren(main.commonTree, true);
  }
  if (!mainFolderName || mainFolderName.toLowerCase().includes('kaizen')) {
    collectChildren(KAIZEN_KNOWLEDGE_BANK_TREE, true);
  }
  return nested;
}


/** Flat rows representing the entire "Common for all vessels" hierarchy */
export function getCommonShipsFlatRows(): import('./types/rows').FlatRow[] {
  const rows: import('./types/rows').FlatRow[] = [];
  let sr = 1;
  const vesselName = 'Common for all vessels';

  for (const main of MAIN_FOLDERS) {
    const group = main.name;
    const commonName = main.commonFolderName ?? COMMON_SHIPS_ROOT;
    const addLeaves = (folder: FolderNode, ancestors: string[]): void => {
      const folderPath = [...ancestors, folder.name];
      if (folder.children?.length) {
        folder.children.forEach(child => addLeaves(child, folderPath));
      } else {
        const category = ancestors[ancestors.length - 1] || folder.name;
        const subCategory = folder.name;
        const subFolderPath = `${group} > ${commonName} > ${folderPath.join(' > ')}`;
        rows.push({
          srNo: String(sr++),
          vesselName,
          group,
          category,
          subCategory,
          subFolderPath,
          fileName: null,
          fileId: null,
          canUpload: true,
          groupKey: `${vesselName}||${group}||${category}||${subCategory}||${subFolderPath}`,
          uploadFolderId: `${group}/${commonName}/${folderPath.join('/')}`,
          monthDriven: false,
        });
      }
    };
    main.commonTree.forEach(folder => addLeaves(folder, []));
  }
  const seen = new Set<string>();
  return rows.filter(row => {
    const key = (row.subFolderPath || '').trim().toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/** Flat rows representing the standalone "Kaizen - Knowledge Bank" hierarchy outside vessels */
export function getKaizenFlatRows(): import('./types/rows').FlatRow[] {
  const rows: import('./types/rows').FlatRow[] = [];
  let sr = 1;
  const vesselName = 'Kaizen - Knowledge Bank';
  const group = 'Kaizen - Knowledge Bank';

  const addLeaves = (folder: FolderNode, ancestors: string[]): void => {
    const folderPath = [...ancestors, folder.name];
    if (folder.children?.length) {
      folder.children.forEach(child => addLeaves(child, folderPath));
    } else {
      const category = ancestors[ancestors.length - 1] || folder.name;
      const subCategory = folder.name;
      const subFolderPath = `Kaizen - Knowledge Bank > ${folderPath.join(' > ')}`;
      rows.push({
        srNo: String(sr++),
        vesselName,
        group,
        category,
        subCategory,
        subFolderPath,
        fileName: null,
        fileId: null,
        canUpload: true,
        groupKey: `${vesselName}||${group}||${category}||${subCategory}||${subFolderPath}`,
        uploadFolderId: `Kaizen - Knowledge Bank/${folderPath.join('/')}`,
        monthDriven: false,
      });
    }
  };

  KAIZEN_KNOWLEDGE_BANK_TREE.forEach(folder => addLeaves(folder, []));
  const seen = new Set<string>();
  return rows.filter(row => {
    const key = (row.subFolderPath || '').trim().toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/** Flat rows representing the entire template hierarchy for a single vessel */
export function getVesselTemplateFlatRows(vesselName: string): import('./types/rows').FlatRow[] {
  const rows: import('./types/rows').FlatRow[] = [];
  let sr = 1;
  for (const main of MAIN_FOLDERS) {
    const group = main.name;
    const addLeaves = (folder: FolderNode, ancestors: string[]): void => {
      const folderPath = [...ancestors, folder.name];
      if (folder.children?.length) {
        folder.children.forEach(child => addLeaves(child, folderPath));
      } else {
        const category = ancestors[ancestors.length - 1] || folder.name;
        const subCategory = folder.name;
        const subFolderPath = `${group} > ${vesselName} > ${folderPath.join(' > ')}`;
        rows.push({
          srNo: String(sr++),
          vesselName,
          group,
          category,
          subCategory,
          subFolderPath,
          fileName: null,
          fileId: null,
          canUpload: true,
          groupKey: `${vesselName}||${group}||${category}||${subCategory}||${subFolderPath}`,
          uploadFolderId: `${group}/${vesselName}/${folderPath.join('/')}`,
          monthDriven: false,
        });
      }
    };
    main.perVesselTree.forEach(folder => addLeaves(folder, []));
  }
  const seen = new Set<string>();
  return rows.filter(row => {
    const key = (row.subFolderPath || '').trim().toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

