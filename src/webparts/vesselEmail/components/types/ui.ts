// UI/domain entity types extracted from VesselEmail.tsx

export interface FormState {
  name: string; imo: string; shipyard: string; hull_number: string; vessel_type: string;
  target_site_ids?: string[];
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
  document_section?: string;
  group?: string;
  category?: string;
  sub_category?: string;
  imo?: string;
  vessel_imo?: string;
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

export interface TagFieldDef {
  key: string;
  label: string;
  type: 'text' | 'textarea' | 'select_vessel' | 'select_dept' | 'select_category' | 'select';
  required: boolean;
  options?: string[] | null;
}

export interface DocumentCategory {
  id: number;
  name: string;
  department: string | null;
  dms_path_template: string | null;
  tag_fields: TagFieldDef[];
  ocr_hints: string[];
  is_active: boolean;
  created_by_email?: string | null;
  created_at?: string | null;
  updated_at?: string | null;
}

export interface OcrFieldResult {
  value: string;
  confidence: number;
  tier: 1 | 2;
}

export type OcrSuggestedTags = Record<string, OcrFieldResult | string | any>;

export interface OcrStagingItem {
  id: number;
  filename: string;
  drive_item_id: string | null;
  source_folder_id: string | null;
  source_subfolder_path: string | null;
  vessel_name: string | null;
  upload_source: 'folder' | 'direct';
  status: 'ocr_pending' | 'ocr_complete' | 'tag_suggested' | 'needs_review' | 'moved' | 'dismissed';
  category_id: number | null;
  category_name: string | null;
  tag_fields: TagFieldDef[];
  suggested_tags: OcrSuggestedTags;
  ocr_text_preview: string | null;
  confidence: number | null;
  matched_keywords: string[];
  final_path: string | null;
  uploaded_by_email: string | null;
  error: string | null;
  created_at: string | null;
  updated_at: string | null;
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
  actionType?: string;
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
    | 'document_deleted'       // file or folder moved to recycle bin
    | 'vessel_unrecognised'   // folder at vessel level not in DMS
    | 'file_outside_structure' // file uploaded outside DMS folder tree
    | 'subfolder_anomaly'      // unexpected subfolder inside a vessel's category
    | 'crud_operation'
    | 'email_alert';
  alert_category?: 'dms' | 'unclassified' | 'classified' | 'crud' | 'email';
  read: boolean;
  read_at?: string | null;
  created_at: string | null;
  updated_at?: string | null;
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
  read?: boolean;
  read_at?: string | null;
  updated_at?: string | null;
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


// ── Vessel Suggestions Module ─────────────────────────────────────────────────

/** Extracted + editable vessel identity extracted from uploaded document file names / OCR text. */
export interface VesselSuggestion {
  /** 7-digit IMO number, e.g. "9812345" */
  imoNumber: string;
  /** Shipyard name, e.g. "Hyundai Heavy Industries" */
  shipyard: string;
  /** Hull / yard number, e.g. "SS268" */
  hullNumber: string;
  /** Vessel type, e.g. "Bulk Carrier" */
  vesselType: string;
  /** Inferred or matched vessel display name */
  vesselName: string;
  /** Non-null when an existing vessel was found in the Term Store that matches */
  matchedExisting: import('./rows').VesselRecord | null;
  /** 0–1 confidence score for the match */
  confidence: number;
  /** File names that contributed to the extraction */
  sourceFiles: string[];
  /** Whether this suggestion was verified from backend OCR extraction */
  ocrVerified?: boolean;
  /** OCR scanned preview text */
  ocrTextPreview?: string;
  /** Backend OCR confidence score (0–1) */
  ocrConfidence?: number;
  /** Keywords matched by OCR */
  matchedKeywords?: string[];
  /** Suggested metadata tags from OCR classification */
  suggestedTags?: Record<string, any>;
  /** Last time this suggestion was detected from scanned/uploaded files */
  lastDetectedAt?: number;
}

/** Upload context captured at the time files were uploaded, used to map the
 * vessel suggestion wizard action back to the correct OCR staging rows. */
export interface VesselSuggestionUploadEntry {
  filename: string;
  drive_item_id: string | null;
  source_subfolder_path?: string | null;
  source_vessel_name?: string | null;
  uploaded_at?: number;
}

/** State shape for the Vessel Suggestions wizard dialog. */
export interface VesselSuggestionDialog {
  open: boolean;
  /** Current wizard step */
  step: 'extract' | 'match' | 'review' | 'done';
  /** Raw File objects being analysed */
  files: File[];
  /** Upload context for robust staging row matching (drive_item_id first). */
  uploadEntries?: VesselSuggestionUploadEntry[];
  /** True while async extraction is running */
  extracting: boolean;
  /** Primary result of extraction + Term Store match */
  suggestion: VesselSuggestion | null;
  /** User-editable copy of suggestion — pre-filled, but all fields are editable */
  editedSuggestion: VesselSuggestion | null;
  /** List of all distinct vessel suggestions found across uploaded files/folder */
  allSuggestions?: VesselSuggestion[];
  /** Currently selected suggestion index from allSuggestions */
  selectedSuggestionIndex?: number;
  /** True while the create-vessel / provision API call is in flight */
  creating: boolean;
  error: string | null;
}

