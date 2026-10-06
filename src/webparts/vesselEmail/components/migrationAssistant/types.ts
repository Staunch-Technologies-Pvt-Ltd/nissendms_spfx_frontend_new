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
  url?: string;
  /** "site_management" = added in Sites → Site Management; "allowed_sites" = .env.migration */
  origin?: 'site_management' | 'allowed_sites';
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

export type S2SItemStatus = 'discovered' | 'copying' | 'copied' | 'metadata_done' | 'metadata_attention' | 'skipped' | 'failed';

export type S2SVerifyStatus = 'ok' | 'changed_by_sharepoint' | 'size_mismatch' | 'hash_mismatch' | 'missing' | 'error';

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
  permissions_report: { principal: string; roles: string[]; status: string; detail: string | null }[] | null;
  verify_status: S2SVerifyStatus | null;
  verify_detail: string | null;
  error: string | null;
}

export type S2SCopyStatus =
  | 'running' | 'paused' | 'cancelled' | 'interrupted' | 'failed'
  | 'completed' | 'completed_with_warnings' | 'completed_with_errors';

export type S2SConflictPolicy = 'skip' | 'replace' | 'rename' | 'fail';

export interface S2SCopyOptions {
  conflictPolicy: S2SConflictPolicy;
  copyPermissions: boolean;
  copyVersions: boolean;
}

export interface S2SVerifySummary {
  checked: number;
  ok: number;
  changed_by_sharepoint: number;
  size_mismatch: number;
  hash_mismatch: number;
  missing: number;
  error: number;
  source_files: number;
  source_folders: number;
  source_bytes: number;
  dest_files_verified: number;
  dest_bytes_verified: number;
  skipped: number;
  not_copied: number;
  passed: boolean;
}

export interface S2SJob {
  id: string;
  status: 'running' | 'done' | 'failed';
  source_site_key: string;
  source_site_label: string;
  source_folder_path: string;
  selected_folders: string[];
  selected_files: string[];
  dest_site_key: string;
  dest_site_label: string;
  dest_folder_path: string;
  total_found: number;
  processed: number;
  started_at: string | null;
  finished_at: string | null;
  confirmed_at: string | null;
  confirmed_by_email: string | null;
  error: string | null;
  copy_status: S2SCopyStatus | null;
  conflict_policy: S2SConflictPolicy | null;
  copy_permissions: boolean;
  copy_versions: boolean;
  copy_started_at: string | null;
  copy_finished_at: string | null;
  copy_summary: (Partial<S2SProgress> & { error?: string }) | null;
  verify_status: 'running' | 'done' | 'failed' | null;
  verified_at: string | null;
  verify_summary: S2SVerifySummary | null;
  live: boolean;
}

export interface S2SProgressEvent {
  seq: number;
  path: string;
  kind: 'file' | 'folder';
  status: 'copied' | 'attention' | 'skipped' | 'failed' | 'created';
  detail: string | null;
}

/** One line of the /stream NDJSON feed. */
export interface S2SProgress {
  type: 'progress';
  job_id: string;
  state: S2SCopyStatus | 'cancelling' | 'verifying' | 'not_started';
  phase: 'folders' | 'files' | 'verifying' | 'done';
  files_total: number;
  files_done: number;
  files_failed: number;
  files_skipped: number;
  folders_total: number;
  folders_done: number;
  folders_failed: number;
  bytes_total: number;
  bytes_done: number;
  metadata_attention: number;
  permissions_attention: number;
  verify_total: number;
  verify_done: number;
  speed_bps: number;
  eta_seconds: number | null;
  elapsed_seconds: number;
  in_flight: string[];
  events: S2SProgressEvent[];
  seq: number;
}

export type S2SStreamLine = S2SProgress | { type: 'done'; job: S2SJob };
