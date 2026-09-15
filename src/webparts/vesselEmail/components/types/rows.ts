// Row & vessel record types extracted from VesselEmail.tsx

export interface FlatRow {
  srNo: string;
  vesselName: string;
  group: string;        // Main folder: Technical & Crewing, Commercial & Chartering, Insurance, Kaizen
  category: string;    // First-level category under vessel: Month End Reports, Agreements, etc.
  subCategory: string; // Leaf folder: Main Engine, Charter Party, etc. (may equal category for single-level)
  subFolderPath: string;
  fileName: string | null;
  fileId: string | null;
  filePending?: boolean;   // true = uploaded & awaiting approval; hides clickable link until approved
  fileUploadedAt?: number; // epoch ms from the file's createdDateTime, used for upload date/time display
  fileSize?: string;
  canUpload: boolean;
  groupKey: string;
  uploadFolderId: string;
  monthDriven: boolean;
}


export interface GroupedRow {
  srNo: string;
  vesselName: string;
  group: string;        // Main folder: Technical & Crewing, Commercial & Chartering, Insurance, Kaizen
  category: string;    // First-level category under vessel: Month End Reports, Agreements, etc.
  subCategory: string; // Leaf folder: Main Engine, Charter Party, etc.
  subFolderPath: string;
  groupKey: string;
  uploadFolderId: string;
  monthDriven: boolean;
  canUpload: boolean;
  files: Array<{ id: string; name: string; size?: string; uploadedAt?: number }>;
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
  is_provisioned?: boolean;
  provisioned_site_ids?: string[];
  site_provisioning_status?: Record<string, string>;
}

