// Row & vessel record types extracted from VesselEmail.tsx

export interface FlatRow {
  srNo: string;
  vesselName: string;
  group: string;
  category: string;
  subFolderPath: string;
  fileName: string | null;
  fileId: string | null;
  filePending?: boolean;   // true = uploaded & awaiting approval; hides clickable link until approved
  canUpload: boolean;
  groupKey: string;
  uploadFolderId: string;
  monthDriven: boolean;
}

export interface GroupedRow {
  srNo: string;
  vesselName: string;
  group: string;
  category: string;
  subFolderPath: string;
  groupKey: string;
  uploadFolderId: string;
  monthDriven: boolean;
  canUpload: boolean;
  files: Array<{ id: string; name: string }>;
}

export interface VesselRecord {
  id: string;
  name: string;
  imo?: string;
  shipyard?: string;
  hull_number?: string;
  vessel_type?: string;
  status?: 'Active' | 'In Maintenance' | 'Inactive' | string;
  image_url?: string;
}
