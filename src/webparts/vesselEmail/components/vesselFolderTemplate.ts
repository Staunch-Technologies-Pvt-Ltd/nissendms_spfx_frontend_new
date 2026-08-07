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
  { name: 'Month End Reports' },
  { name: 'Service Agreements' },
  { name: 'Registration' },
  {
    name: 'Drawings and Manuals',
    children: [
      {
        name: 'Drawings',
        children: [
          { name: 'Archive' },
          { name: 'Basic' },
          { name: 'Electrical' },
          { name: 'Engine' },
          { name: 'Hull' },
          { name: 'Other Drawings' },
          { name: 'Safety' },
        ],
      },
      {
        name: 'Manuals',
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
    ],
  },
  { name: 'Incidents' },
  { name: 'Crewing' },
];

const TECHNICAL_CREWING_COMMON: FolderNode[] = [
  {
    name: 'Common (for all ships)',
    children: [
      { name: 'Vendor & Service Agreements' },
      { name: 'Vendor Management' },
    ],
  },
];

// ── Folder-2: Commercial and Chartering ──────────────────────────────────────

const COMMERCIAL_CHARTERING_PER_VESSEL: FolderNode[] = [
  { name: 'Agreements' },
  { name: 'Invoices & Payments' },
  { name: 'Claims & Disputes' },
  { name: 'To be Classified' },
];

const COMMERCIAL_CHARTERING_COMMON: FolderNode[] = [
  {
    name: 'Common Agreements (Not Ship Specific)',
    children: [
      { name: 'Agreements' },
      { name: 'To be Classified' },
    ],
  },
];

// ── Folder-3: Insurance ───────────────────────────────────────────────────────

const INSURANCE_PER_VESSEL: FolderNode[] = [
  { name: 'P&I' },
  { name: 'H&M' },
  { name: 'War Risk' },
  { name: 'Flag and MPA' },
];

const INSURANCE_COMMON: FolderNode[] = [
  {
    name: 'Common (Not Ship Specific)',
    children: [
      { name: 'Agreements' },
      { name: 'Miscellaneous' },
    ],
  },
];

// ── Kaizen - Knowledge Bank ───────────────────────────────────────────────────

const KAIZEN_KNOWLEDGE_BANK_PER_VESSEL: FolderNode[] = [];

const KAIZEN_KNOWLEDGE_BANK_COMMON: FolderNode[] = [
  { name: 'Templates' },
  { name: 'Procedures and Work Instructions' },
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

export const VESSEL_MANAGEMENT_ROOT = 'Vessel Management';

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
  {
    name: 'Kaizen - Knowledge Bank',
    perVesselTree: KAIZEN_KNOWLEDGE_BANK_PER_VESSEL,
    commonTree: KAIZEN_KNOWLEDGE_BANK_COMMON,
  },
];
