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
exports.__esModule = true;
exports.MAIN_FOLDERS = exports.VESSEL_MANAGEMENT_ROOT = void 0;
// ── Folder-1: Technical & Crewing ────────────────────────────────────────────
var TECHNICAL_CREWING_PER_VESSEL = [
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
                ]
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
                ]
            },
        ]
    },
    { name: 'Incidents' },
    { name: 'Crewing' },
];
var TECHNICAL_CREWING_COMMON = [
    {
        name: 'Common (for all ships)',
        children: [
            { name: 'Vendor & Service Agreements' },
            { name: 'Vendor Management' },
        ]
    },
];
// ── Folder-2: Commercial and Chartering ──────────────────────────────────────
var COMMERCIAL_CHARTERING_PER_VESSEL = [
    { name: 'Agreements' },
    { name: 'Invoices & Payments' },
    { name: 'Claims & Disputes' },
    { name: 'To be Classified' },
];
var COMMERCIAL_CHARTERING_COMMON = [
    {
        name: 'Common Agreements (Not Ship Specific)',
        children: [
            { name: 'Agreements' },
            { name: 'To be Classified' },
        ]
    },
];
// ── Folder-3: Insurance ───────────────────────────────────────────────────────
var INSURANCE_PER_VESSEL = [
    { name: 'P&I' },
    { name: 'H&M' },
    { name: 'War Risk' },
    { name: 'Flag and MPA' },
];
var INSURANCE_COMMON = [
    {
        name: 'Common (Not Ship Specific)',
        children: [
            { name: 'Agreements' },
            { name: 'Miscellaneous' },
        ]
    },
];
var KAIZEN_KNOWLEDGE_BANK_PER_VESSEL = [];
var KAIZEN_KNOWLEDGE_BANK_COMMON = [
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
            { name: 'Shipyard' }
        ]
    }
];
// ── Exported template ─────────────────────────────────────────────────────────
exports.VESSEL_MANAGEMENT_ROOT = 'Vessel Management';
exports.MAIN_FOLDERS = [
    {
        name: 'Technical & Crewing',
        perVesselTree: TECHNICAL_CREWING_PER_VESSEL,
        commonTree: TECHNICAL_CREWING_COMMON
    },
    {
        name: 'Commercial & Chartering',
        perVesselTree: COMMERCIAL_CHARTERING_PER_VESSEL,
        commonTree: COMMERCIAL_CHARTERING_COMMON
    },
    {
        name: 'Insurance',
        perVesselTree: INSURANCE_PER_VESSEL,
        commonTree: INSURANCE_COMMON
    },
    {
        name: 'Kaizen - Knowledge Bank',
        perVesselTree: KAIZEN_KNOWLEDGE_BANK_PER_VESSEL,
        commonTree: KAIZEN_KNOWLEDGE_BANK_COMMON
    }
];
