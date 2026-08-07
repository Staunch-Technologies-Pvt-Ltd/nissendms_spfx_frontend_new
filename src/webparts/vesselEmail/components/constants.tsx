// Helper functions, lookup tables, and initial mock data extracted from VesselEmail.tsx
import * as React from 'react';
import type { DocumentItem, TemplateItem, ApprovalItem, NotificationItem, UserItem } from './types/ui';

export function cleanName(name: string): string {
  return (name || '').replace(/^_+|_+$/g, '').trim();
}

export const VESSEL_TYPES = [
  'Bulk Carrier', 'Container Ship', 'Gas Carrier',
  'Oil Tanker', 'Chemical Tanker', 'General Cargo', 'Offshore Support', 'Other Cargo Ships',
];

export const DATASOURCE_TAGS_MAP: Record<string, string> = {
  contract: 'CP',
  other_contract: 'Other Contracts',
  vessels_certificate: 'Certificate',
  vessels_drawing: 'Drawing',
  mail: 'Mail',
  imo: 'IMO',
  uscg: 'USCG',
  msib: 'USCG MSIB',
  imo_flag_country_others: 'IMO Flag Country Others',
  panama_flag_circular: 'Panama Flag Circular',
  imo_flag_country_flag: 'Flag',
  nk: 'NK',
  japan_p_and_i: 'Japan P&I',
  ukpandi: 'UK P&I',
  gard: 'GARD',
  scmg: 'Standard Club',
  britannia_p_and_i: 'Britannia P&I',
  bimco: 'BIMCO',
  security_information: 'Security Information',
  omc_kaikoumu: 'OMC Marine & Tech. Support Center',
  ice_information: 'Ice Information',
  right_ship: 'RightShip',
  others: 'Others',
};

export function suggestTagFromFilename(filename: string): string {
  const s = (filename || '').toLowerCase();
  if (['cp', 'charter', 'contract', 'agreement', 'party'].some(k => s.indexOf(k) !== -1)) return 'contract';
  if (['vendor', 'supplier', 'subcontract', 'other_contract'].some(k => s.indexOf(k) !== -1)) return 'other_contract';
  if (['cert', 'certificate', 'class', 'survey', 'statutory', 'audit'].some(k => s.indexOf(k) !== -1)) return 'vessels_certificate';
  if (['draw', 'plan', 'schematic', 'manual', 'diagram', 'blueprint', 'ga_plan'].some(k => s.indexOf(k) !== -1)) return 'vessels_drawing';
  if (s.indexOf('msib') !== -1) return 'msib';
  if (['uscg', 'coastguard', 'coast guard'].some(k => s.indexOf(k) !== -1)) return 'uscg';
  if (['panama', 'circular'].some(k => s.indexOf(k) !== -1)) return 'panama_flag_circular';
  if (['nk', 'nippon'].some(k => s.indexOf(k) !== -1)) return 'nk';
  if (['ukpandi', 'uk_p_and_i', 'uk p&i'].some(k => s.indexOf(k) !== -1)) return 'ukpandi';
  if (s.indexOf('gard') !== -1) return 'gard';
  if (['scmg', 'standard_club', 'standard club'].some(k => s.indexOf(k) !== -1)) return 'scmg';
  if (s.indexOf('britannia') !== -1) return 'britannia_p_and_i';
  if (['pandi', 'p&i', 'pi', 'japan_p_and_i', 'protection'].some(k => s.indexOf(k) !== -1)) return 'japan_p_and_i';
  if (s.indexOf('bimco') !== -1) return 'bimco';
  if (['security', 'isps', 'sec_info'].some(k => s.indexOf(k) !== -1)) return 'security_information';
  if (['kaikoumu', 'omc', 'tech_support'].some(k => s.indexOf(k) !== -1)) return 'omc_kaikoumu';
  if (['ice', 'ice_info'].some(k => s.indexOf(k) !== -1)) return 'ice_information';
  if (['rightship', 'right_ship', 'right ship'].some(k => s.indexOf(k) !== -1)) return 'right_ship';
  if (['flag', 'flag_state'].some(k => s.indexOf(k) !== -1)) return 'imo_flag_country_flag';
  if (s.indexOf('imo') !== -1) return 'imo';
  if (['mail', 'email', 'letter', 'memo', 'msg', 'eml'].some(k => s.indexOf(k) !== -1)) return 'mail';
  return 'mail';
}

export function badge(color: string, text: string): React.ReactElement {
  const map: Record<string, [string, string]> = {
    blue: ['#e1efff', '#0078d4'], orange: ['#fff4ce', '#8a5700'],
    purple: ['#ede8f5', '#5c2d91'], green: ['#dff6dd', '#107c10'],
    red: ['#fde7e9', '#a4262c'], default: ['#f3f2f1', '#323130'],
  };
  const pair = map[color] || map['default'];
  return <span style={{ display: 'inline-block', borderRadius: 12, padding: '3px 10px', fontSize: 11, fontWeight: 600, background: pair[0], color: pair[1], whiteSpace: 'nowrap' }}>{text}</span>;
}

export const GROUP_COLORS: Record<string, string> = {
  Agreements: 'blue', 'Claims & Disputes': 'orange', Crewing: 'blue',
  Electrical: 'orange', Hull: 'blue', Incidents: 'red', Insurance: 'orange',
  Machinery: 'purple', Safety: 'green', 'Technical & Crewing': 'orange',
  'Commercial & Chartering': 'green', 'To be Classified': 'default',
};

// ── Default Mock Data ────────────────────────────────────────────────────────

export const INITIAL_MOCK_DOCUMENTS: DocumentItem[] = [
  { id: 'd1', name: 'Certificates', vessel: '-', type: 'Folder', expiryDate: '-', status: 'Valid', modified: 'May 12, 2024', isFolder: true },
  { id: 'd2', name: 'Crew Documents', vessel: '-', type: 'Folder', expiryDate: '-', status: 'Valid', modified: 'May 11, 2024', isFolder: true },
  { id: 'd3', name: 'Insurance', vessel: '-', type: 'Folder', expiryDate: '-', status: 'Valid', modified: 'May 10, 2024', isFolder: true },
  { id: 'd4', name: 'Insurance Certificate.pdf', vessel: 'Ocean Star', type: 'Insurance', expiryDate: 'May 25, 2024', status: 'Expiring Soon', modified: 'May 12, 2024' },
  { id: 'd5', name: 'Crew List.docx', vessel: 'Ocean Star', type: 'Crew', expiryDate: 'Jun 10, 2024', status: 'Valid', modified: 'May 11, 2024' },
  { id: 'd6', name: 'Maintenance Log.xlsx', vessel: 'Sea Breeze', type: 'Maintenance', expiryDate: 'May 20, 2024', status: 'Valid', modified: 'May 10, 2024' },
  { id: 'd7', name: 'Safety Certificate.pdf', vessel: 'Blue Horizon', type: 'Certificate', expiryDate: 'Apr 15, 2024', status: 'Expired', modified: 'May 09, 2024' },
];

export const INITIAL_MOCK_TEMPLATES: TemplateItem[] = [
  { id: 't1', name: 'Insurance Certificate Template.docx', type: 'Insurance', description: 'Template for Insurance Certificate', modified: 'May 12, 2024' },
  { id: 't2', name: 'Crew List Template.docx', type: 'Crew', description: 'Template for Crew List', modified: 'May 11, 2024' },
  { id: 't3', name: 'Maintenance Report Template.docx', type: 'Maintenance', description: 'Template for Maintenance Report', modified: 'May 10, 2024' },
  { id: 't4', name: 'Safety Certificate Template.docx', type: 'Certificate', description: 'Template for Safety Certificate', modified: 'May 09, 2024' },
  { id: 't5', name: 'Survey Report Template.docx', type: 'Survey', description: 'Template for Survey Report', modified: 'May 08, 2024' },
];

export const INITIAL_MOCK_APPROVALS: ApprovalItem[] = [
  { id: 'a1', documentName: 'Insurance Certificate.pdf', vessel: 'Ocean Star', requestedBy: 'John Doe', requestedOn: 'May 12, 2024', status: 'Pending' },
  { id: 'a2', documentName: 'Crew List.docx', vessel: 'Ocean Star', requestedBy: 'Priya Sharma', requestedOn: 'May 11, 2024', status: 'Pending' },
  { id: 'a3', documentName: 'Maintenance Log.xlsx', vessel: 'Sea Breeze', requestedBy: 'Rohit Kumar', requestedOn: 'May 10, 2024', status: 'Pending' },
  { id: 'a4', documentName: 'Safety Certificate.pdf', vessel: 'Blue Horizon', requestedBy: 'Ankita Verma', requestedOn: 'May 09, 2024', status: 'Pending' },
  { id: 'a5', documentName: 'Survey Report.pdf', vessel: 'Pacific Dawn', requestedBy: 'Vikram Singh', requestedOn: 'May 08, 2024', status: 'Pending' },
  { id: 'a6', documentName: 'Cargo Manifest.pdf', vessel: 'Golden Pearl', requestedBy: 'John Doe', requestedOn: 'May 05, 2024', status: 'Approved' },
  { id: 'a7', documentName: 'De-ballasting Plan.docx', vessel: 'Atlantic Wave', requestedBy: 'Rohit Kumar', requestedOn: 'May 02, 2024', status: 'Rejected' },
];

export const INITIAL_MOCK_NOTIFICATIONS: NotificationItem[] = [
  { id: 'n1', title: 'Document Expired', message: 'Document expired: Safety Certificate.pdf for Blue Horizon', timestamp: 'May 12, 2024 10:30 AM', priority: 'High', read: false, type: 'alert' },
  { id: 'n2', title: 'Document Expiring Soon', message: 'Document expiring in 7 days: Insurance Certificate.pdf for Ocean Star', timestamp: 'May 12, 2024 09:15 AM', priority: 'Medium', read: false, type: 'warning' },
  { id: 'n3', title: 'Document Approved', message: 'Document approved: Crew List.docx for Ocean Star', timestamp: 'May 11, 2024 04:20 PM', priority: 'Low', read: true, type: 'success' },
  { id: 'n4', title: 'New Document Uploaded', message: 'New document uploaded: Maintenance Log.xlsx for Sea Breeze', timestamp: 'May 11, 2024 11:05 AM', priority: 'Low', read: true, type: 'info' },
  { id: 'n5', title: 'Approval Requested', message: 'Approval requested: Survey Report.pdf for Pacific Dawn', timestamp: 'May 10, 2024 02:45 PM', priority: 'Medium', read: false, type: 'warning' },
];

export const INITIAL_MOCK_USERS: UserItem[] = [
  { id: 'u1', name: 'Priya Sharma', email: 'priya.sharma@company.com', role: 'Administrator', status: 'Active', lastLogin: 'May 12, 2024' },
  { id: 'u2', name: 'John Doe', email: 'john.doe@company.com', role: 'Manager', status: 'Active', lastLogin: 'May 12, 2024' },
  { id: 'u3', name: 'Rohit Kumar', email: 'rohit.kumar@company.com', role: 'User', status: 'Active', lastLogin: 'May 11, 2024' },
  { id: 'u4', name: 'Ankita Verma', email: 'ankita.verma@company.com', role: 'User', status: 'Active', lastLogin: 'May 10, 2024' },
  { id: 'u5', name: 'Vikram Singh', email: 'vikram.singh@company.com', role: 'Reviewer', status: 'Inactive', lastLogin: 'May 08, 2024' },
];
