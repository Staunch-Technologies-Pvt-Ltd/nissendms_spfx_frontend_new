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
      { name: 'Drawing' },
      { name: 'Manual' },
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

// ── Kaizen - Knowledge Bank ───────────────────────────────────────────────────

// Every vessel has the same four top-level main folders.  Kaizen uses the
// same category structure for a vessel as it does in the common area.
const KAIZEN_KNOWLEDGE_BANK_PER_VESSEL: FolderNode[] = [
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

const KAIZEN_KNOWLEDGE_BANK_COMMON: FolderNode[] = [
  { name: 'Templates' },
  {
    name: 'Procedures and Work Instructions',
    children: [
      { name: 'Equipment Maker' },
      { name: 'Class' },
      { name: 'Flag ⁄ Port State' },
      { name: 'SIRE⁄OCIMF⁄RightShip' },
      { name: 'Shipyard' },
    ],
  },
  { name: 'Lessons Learned' },
  {
    name: 'Circulars and Guidance',
    children: [
      { name: 'Equipment Maker' },
      { name: 'Class' },
      { name: 'Flag ⁄ Port State' },
      { name: 'SIRE⁄OCIMF⁄RightShip' },
      { name: 'Shipyard' },
    ],
  },
];

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
