// Types for the Migration Assistant module — ported from the standalone
// SharePoint AI Migration Assistant project's frontend/src/api.ts, trimmed
// to what this module's UI uses. Kept in its own file (no dependency on the
// host app's rows/ui types) so this module stays a self-contained package.

export type MigrationStatus =
  | 'discovered'
  | 'extracting'
  | 'suggested'
  | 'needs_review'
  | 'failed'
  | 'moved';

export interface MigrationItem {
  id: string;
  job_id: string;
  source_path: string;
  filename: string;
  content_type: string;
  size: number;
  status: MigrationStatus;
  extracted_text_excerpt: string | null;
  suggested_path: string | null;
  category: string | null;
  subcategory: string | null;
  confidence: number | null;
  reason: string | null;
  keywords: string[];
  detected_vessel_name: string | null;
  vessel_exists: boolean | null;
  vessel_confidence: number | null;
  term_tag_status: 'applied' | 'partial' | 'unmapped' | 'skipped' | 'error' | null;
  term_tag_report: { field: string; status: string; detail: string }[];
  overridden: boolean;
  final_path: string | null;
  decided_by_email: string | null;
  decided_at: string | null;
  error: string | null;
  created_at: string | null;
  updated_at: string | null;
}

export interface MigrationScanJob {
  id: string;
  status: 'running' | 'done' | 'failed';
  source_folder: string;
  subfolders: string[];
  files: string[];
  vessel_name: string;
  vessel_path: string;
  auto_detect_vessel: boolean;
  total_found: number;
  processed: number;
  started_at: string | null;
  finished_at: string | null;
  confirmed_at: string | null;
  confirmed_by_email: string | null;
  error: string | null;
}

export interface HierarchyNode {
  id: string;
  name: string;
  path: string;
  children: HierarchyNode[];
}

export interface SourceFolder {
  id: string;
  name: string;
  path: string;
}

export interface SourceFile {
  id: string;
  name: string;
  size: number;
  content_type: string;
}

export interface ConfirmSummary {
  total: number;
  categorized: number;
  moved: number;
  failed: number;
  failed_items: { filename: string; error: string }[];
  vessels_created: number;
  tags_applied: number;
  tags_unmapped: number;
  tags_failed: number;
}

export interface ExistingFileTagRow {
  id: string;
  list_item_id: number;
  filename: string;
  folder_path: string;
  current: { group: string | null; category: string | null; vessel: string | null };
  derived: { group: string; category: string; vessel: string } | null;
  terms: Record<
    string,
    { label: string; term_label: string | null; guid: string | null; resolved: boolean; reason: string }
  >;
  ready: boolean;
  category_skipped: boolean;
  error: string | null;
}

export interface ExistingFileTagScan {
  root_path: string;
  bindings: Record<string, { field: string; term_set_id: string }>;
  files: ExistingFileTagRow[];
  summary: {
    total_files: number;
    ready_files: number;
    fully_tagged: number;
    category_skipped: number;
    missing_taxonomy: number;
    missing_terms: number;
  };
}

export interface ExistingFileTagApply {
  total_selected: number;
  successful: number;
  failed: number;
  fully_tagged: number;
  category_skipped: number;
  results: Array<ExistingFileTagRow & { success: boolean; failure_reason: string | null }>;
}

export interface TaggedScanJob {
  id: string;
  root_path: string;
  status: 'ready' | 'failed';
  summary: {
    root_path: string;
    total_files: number;
    ready_files: number;
    fully_tagged: number;
    category_skipped: number;
    missing_taxonomy: number;
    missing_terms: number;
    selected_count?: number;
    selected_ready?: number;
  };
  created_at: string | null;
}

export interface S2SSite {
  key: string;
  label: string;
}

export interface S2SDrive {
  id: string;
  name: string;
}

export interface S2SFolder {
  id: string;
  name: string;
  path: string;
}

export interface S2SFile {
  id: string;
  name: string;
  size: number;
  content_type: string;
}

export type S2SItemStatus = 'discovered' | 'copying' | 'copied' | 'metadata_done' | 'metadata_attention' | 'failed';

export interface S2SItem {
  id: string;
  job_id: string;
  kind: 'folder' | 'file';
  relative_path: string;
  name: string;
  size: number;
  status: S2SItemStatus;
  dest_item_id: string | null;
  metadata_report: { field: string; kind: string; status: string; detail: string }[] | null;
  error: string | null;
}

export interface S2SJob {
  id: string;
  status: 'running' | 'done' | 'failed';
  source_site_key: string;
  source_folder_path: string;
  selected_folders: string[];
  selected_files: string[];
  dest_site_key: string;
  dest_folder_path: string;
  total_found: number;
  processed: number;
  started_at: string | null;
  finished_at: string | null;
  confirmed_at: string | null;
  confirmed_by_email: string | null;
  error: string | null;
}

export interface S2SConfirmSummary {
  total: number;
  folders_created: number;
  folders_existing: number;
  folders_failed: number;
  files_copied: number;
  files_failed: number;
  metadata_migrated: number;
  metadata_attention: number;
  managed_metadata_applied: number;
  managed_metadata_unmapped: number;
  failed_items: { path: string; error: string }[];
  status: 'completed' | 'completed_with_warnings' | 'completed_with_errors';
}
