// UI/domain entity types extracted from VesselEmail.tsx

export interface FormState {
  name: string; imo: string; shipyard: string; hull_number: string; vessel_type: string;
}

export interface DocPreviewItem {
  fileName: string;
  fileId?: string | null;
  vesselName: string;
  group: string;
  category: string;
  folderPath: string;
  size?: string;
  date?: string;
}

export interface DeletedNode {
  id: string;
  name: string;
  kind: string;
  original_path?: string;
  main_folder?: string;
  deleted_at?: string | null;
  archived_at?: string | null;
  item_type?: string;
  size?: number | null;
  modified?: string | null;
  ext?: string;
  imo?: string;
  shipyard?: string;
  hull_number?: string;
  vessel_type?: string;
}

export interface DocumentItem {
  id: string;
  name: string;
  vessel: string;
  type: string;
  expiryDate: string;
  status: 'Valid' | 'Expiring Soon' | 'Expired';
  modified: string;
  isFolder?: boolean;
}

export interface TemplateItem {
  id: string;
  name: string;
  type: string;
  description: string;
  modified: string;
}

export interface ApprovalItem {
  id: string;
  documentName: string;
  vessel: string;
  requestedBy: string;
  requestedOn: string;
  status: 'Pending' | 'Approved' | 'Rejected';
}

export interface NotificationItem {
  id: string;
  title: string;
  message: string;
  timestamp: string;
  priority: 'High' | 'Medium' | 'Low';
  read: boolean;
  type: 'warning' | 'info' | 'success' | 'alert';
}

export interface UserItem {
  id: string;
  name: string;
  email: string;
  role: 'Administrator' | 'Manager' | 'User' | 'Reviewer';
  status: 'Active' | 'Inactive';
  lastLogin: string;
}

export interface FolderAnomalyItem {
  id: number;
  drive_item_id: string;
  name: string;
  item_type: 'folder' | 'file';
  anomaly_type: 'vessel_level_unmatched' | 'main_folder_unmatched' | 'subfolder_unmatched' | 'classified_normal';
  department: string;
  vessel_name: string | null;
  spo_path: string;
  resolved: boolean;
  detected_at: string | null;
}

export interface NormalFolderRecord {
  id: number | null;
  drive_item_id: string;
  name: string;
  item_type: 'folder' | 'file';
  spo_path: string;
  department: string;
  vessel_name: string | null;
  detected_at: string | null;
}


