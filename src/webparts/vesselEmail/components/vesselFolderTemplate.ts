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
];

// ── Folder-3: Insurance ───────────────────────────────────────────────────────

const INSURANCE_PER_VESSEL: FolderNode[] = [
  { name: 'P&I' },
  { name: 'H&M' },
  { name: 'War Risk' },
  { name: 'Flag / MPA' },
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
      { name: 'Flag / Port State' },
      { name: 'SIRE/OCIMF/RightShip' },
      { name: 'Shipyard' },
    ],
  },
  { name: 'Lessons Learned' },
  {
    name: 'Circulars and Guidance',
    children: [
      { name: 'Equipment Maker' },
      { name: 'Class' },
      { name: 'Flag / Port State' },
      { name: 'SIRE/OCIMF/RightShip' },
      { name: 'Shipyard' },
    ],
  },
];

export const KAIZEN_KNOWLEDGE_BANK_COMMON = KAIZEN_KNOWLEDGE_BANK_TREE;

// ── Exported template ─────────────────────────────────────────────────────────

// Root container under Documents
export const VESSEL_MANAGEMENT_ROOT = 'Vessels';
// Specific vessels go under Vessels/Specific Vessels/{VesselName}
export const SPECIFIC_VESSELS_ROOT = 'Specific Vessels';
// Common folders go under Vessels/Common for all ships
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
  },
  {
    name: 'Insurance',
    perVesselTree: INSURANCE_PER_VESSEL,
    commonTree: INSURANCE_COMMON,
  },
];

/** Display-ready folder lookups used by Documents and Folder view. */
export function folderNamesByMainFolder(common: boolean = false): Record<string, string[]> {
  const map = Object.fromEntries(MAIN_FOLDERS.map(main => [
    main.name,
    (common ? main.commonTree : main.perVesselTree).map(folder => folder.name),
  ]));
  map['Kaizen - Knowledge Bank'] = KAIZEN_KNOWLEDGE_BANK_TREE.map(f => f.name);
  return map;
}

export function subfolderNamesByFolder(common: boolean = false): Record<string, string[]> {
  const result: Record<string, string[]> = {};
  const visit = (folders: FolderNode[]): void => {
    for (const folder of folders) {
      if (folder.children?.length) {
        result[folder.name] = folder.children.map(child => child.name);
        visit(folder.children);
      }
    }
  };
  for (const main of MAIN_FOLDERS) visit(common ? main.commonTree : main.perVesselTree);
  visit(KAIZEN_KNOWLEDGE_BANK_TREE);
  return result;
}

/** Flat rows representing the entire "Common for all vessels" hierarchy */
export function getCommonShipsFlatRows(): import('./types/rows').FlatRow[] {
  const rows: import('./types/rows').FlatRow[] = [];
  let sr = 1;
  const vesselName = 'Common for all vessels';

  for (const main of MAIN_FOLDERS) {
    const group = main.name;
    const addLeaves = (folder: FolderNode, ancestors: string[]): void => {
      const folderPath = [...ancestors, folder.name];
      if (folder.children?.length) {
        folder.children.forEach(child => addLeaves(child, folderPath));
      } else {
        const category = ancestors[ancestors.length - 1] || folder.name;
        const subCategory = folder.name;
        const subFolderPath = `Common for all ships > ${group} > ${folderPath.join(' > ')}`;
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
          uploadFolderId: `Vessels/Common for all ships/${group}/${folderPath.join('/')}`,
          monthDriven: false,
        });
      }
    };
    main.commonTree.forEach(folder => addLeaves(folder, []));
  }
  return rows;
}

/** Flat rows representing the standalone "Kaizen - Knowledge Bank" hierarchy outside vessels */
export function getKaizenFlatRows(): import('./types/rows').FlatRow[] {
  const rows: import('./types/rows').FlatRow[] = [];
  let sr = 1;
  const vesselName = 'Kaizen - Knowledge Bank';
  const group = 'Kaizen - Knowledge Bank';

  for (const folder of KAIZEN_KNOWLEDGE_BANK_TREE) {
    if (folder.children && folder.children.length > 0) {
      for (const child of folder.children) {
        const category = folder.name;
        const subCategory = child.name;
        const subFolderPath = `Kaizen - Knowledge Bank > ${category} > ${subCategory}`;
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
          uploadFolderId: `Kaizen - Knowledge Bank/${category}/${subCategory}`,
          monthDriven: false,
        });
      }
    } else {
      const category = folder.name;
      const subCategory = folder.name;
      const subFolderPath = `Kaizen - Knowledge Bank > ${category}`;
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
        uploadFolderId: `Kaizen - Knowledge Bank/${category}`,
        monthDriven: false,
      });
    }
  }
  return rows;
}
