"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.MAIN_FOLDERS = exports.VESSEL_MANAGEMENT_ROOT = void 0;
// ── Folder-1: Technical & Crewing ────────────────────────────────────────────
var TECHNICAL_CREWING_PER_VESSEL = [
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
            { name: 'Purchase Order and Vendor Invoice' },
        ],
    },
    { name: 'Incidents' },
    { name: 'Crewing' },
    { name: 'To be Classified' },
];
var TECHNICAL_CREWING_COMMON = [
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
var COMMERCIAL_CHARTERING_PER_VESSEL = [
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
var COMMERCIAL_CHARTERING_COMMON = [
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
var INSURANCE_PER_VESSEL = [
    { name: 'P&I' },
    { name: 'H&M' },
    { name: 'War Risk' },
    { name: 'Flag / MPA' },
    { name: 'USA Related' },
];
var INSURANCE_COMMON = [
    { name: 'Agreements' },
    { name: 'Miscellaneous' },
];
// ── Kaizen - Knowledge Bank ───────────────────────────────────────────────────
// Every vessel has the same four top-level main folders.  Kaizen uses the
// same category structure for a vessel as it does in the common area.
var KAIZEN_KNOWLEDGE_BANK_PER_VESSEL = [
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
var KAIZEN_KNOWLEDGE_BANK_COMMON = [
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
exports.VESSEL_MANAGEMENT_ROOT = '';
exports.MAIN_FOLDERS = [
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
