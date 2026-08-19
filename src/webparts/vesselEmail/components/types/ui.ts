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
  vessel_name?: string;
  category?: string;
  sub_category?: string;
  imo?: string;
  shipyard?: string;
  hull_number?: string;
  vessel_type?: string;
  /** True when the item was soft-deleted via Graph and now lives in the SPO site recycle bin. */
  in_spo_recycle_bin?: boolean;
  /** The SPO recycle bin item ID (different from the drive item ID) — populated after soft-delete.
   *  Used for permanent deletion via DELETE /sites/{siteId}/recycleBin/{recycleBinItemId}. */
  recycle_bin_item_id?: string;
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

/** A folder-creation alert surfaced by the top-header alert bell. Covers newly
 *  created SharePoint Online folders (manual sub-folders), newly provisioned
 *  vessels, and SPO anomaly detections (unrecognised folders/files). */
export interface AlertItem {
  id: string;
  drive_item_id: string | null;
  folder_name: string;
  folder_path: string;
  parent_folder_id: string | null;
  vessel_name: string | null;
  department: string;
  created_by_email: string;
  created_by_name: string;
  alert_type:
    | 'folder_created'
    | 'vessel_provisioned'
    | 'vessel_deleted'         // vessel moved to recycle bin
    | 'vessel_unrecognised'   // folder at vessel level not in DMS
    | 'file_outside_structure' // file uploaded outside DMS folder tree
    | 'subfolder_anomaly'      // unexpected subfolder inside a vessel's category
    | 'crud_operation'
    | 'email_alert';
  alert_category?: 'dms' | 'crud' | 'email';
  read: boolean;
  created_at: string | null;
  // Anomaly-specific fields (populated for anomaly alert types)
  anomaly_id?: number;
  item_type?: 'folder' | 'file';
  spo_path?: string;
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

/** Dialog shown when a file is uploaded directly to Specific Vessels/ or any
 *  sub-folder path outside the DMS structure. */
export interface SpoFileAlertDialog {
  fileId: string;
  fileName: string;
  spoPath: string;
  vesselName: string | null;
  /** Available DMS sub-folder rows for the detected vessel (for move target). */
  subFolderOptions: Array<{ label: string; groupKey: string; uploadFolderId: string; subFolderPath: string }>;
  moving: boolean;
  moved: boolean;
  error: string | null;
}


