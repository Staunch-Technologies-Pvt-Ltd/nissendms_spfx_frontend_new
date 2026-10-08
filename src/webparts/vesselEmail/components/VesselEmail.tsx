import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IVesselEmailProps } from './IVesselEmailProps';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from './vesselImagePool';
import { MAIN_FOLDERS } from './vesselFolderTemplate';
import { clay, applyColorTheme, DEFAULT_CLAY_COLORS } from './clayTheme';
import type { ClayColorTheme } from './clayTheme';

/** Per-folder outcome of a provisioning attempt. Folder provisioning now
 * always goes through the backend's flat-root `/provision-sites` endpoint
 * (a single folder, single API call) rather than the old client-side
 * multi-folder Graph builder (graphFolderService.ts, retired), so this list
 * is normally empty on success — it's kept only so the existing
 * folderCreationResults/folderCreationFeed state and the vessel-provisioning
 * dialogs that read them keep compiling and can still surface a per-item
 * failure list on the rare path that reports one. */
export interface FolderResult {
  path: string;
  id?: string;
  status: 'created' | 'existed' | 'failed';
  error?: string;
}
import {
  createSyncScheduler, SyncScheduler, DeltaSyncResult,
  mergeNodeIntoMap, removeNodeFromMap, SpoFolderNode, fetchFolderChildren, repathDescendants } from './deltaSync';



import type { FlatRow, GroupedRow, VesselRecord } from './types/rows';
import type { BentoEmailLog } from './types/bento';
import type { AppView, ModalMode } from './types/view';
import type {
  FormState, DocPreviewItem, DeletedNode, DocumentItem, TemplateItem,
  UserItem, UserSitePermissionItem, FolderAnomalyItem, NormalFolderRecord, AlertItem,
  ApprovalItem, VesselSuggestion, VesselSuggestionDialog, OcrStagingItem, VesselSuggestionUploadEntry,
} from './types/ui';


import {
  cleanName, INITIAL_MOCK_APPROVALS, INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES,
} from './constants';
import { renderSidebar } from './pages/Sidebar';
import { renderLayout } from './pages/AppLayout';
import { renderDocPreviewDrawer } from './pages/DocPreviewDrawer';
import { CopilotSearchPanel } from './copilot/CopilotSearchPanel';
import { renderDashboard, DashboardStats } from './pages/DashboardPage';
import { renderDocumentsPage } from './pages/DocumentsPage';
import { SitesPage } from './pages/SitesPage';
import { renderVesselsPage, renderClassifyDialog } from './pages/VesselsPage';
import { renderUsersPage } from './pages/UsersPage';
import { renderSettingsPage } from './pages/SettingsPage';
import { renderProfilePage } from './pages/ProfilePage';
import { renderAuthPage, AuthPageMode } from './pages/AuthPage';
import { renderBentoEmailDashboardPage } from './pages/BentoEmailDashboardPage';
import { renderRecycleBinPage } from './pages/RecycleBinPage';
import { ArchivePage } from './pages/ArchivePage';
import { renderAlertsPage } from './pages/AlertsPage';
import { ComposeMailModal } from './mail/ComposeMailModal';
import { renderVesselForm } from './modals/VesselFormModal';
import { renderDeleteModal } from './modals/DeleteVesselModal';
import { renderArchivePickerModal, ArchivePickerDialogState } from './modals/ArchiveSelectionModal';
import { renderCreateFolderModal, CreateFolderDialogState } from './modals/CreateFolderModal';
import { BulkUploadModal, BulkUploadFile, extractFilesFromDataTransfer } from './BulkUploadModal';
import { renderVesselSuggestionsModal } from './pages/VesselSuggestionsModal';
import { isMobileWidth, isTabletOrBelow } from './responsive';
import { MigrationAssistantModule } from './migrationAssistant/MigrationAssistantModule';

// ── Error Boundary ────────────────────────────────────────────────────────────
// Prevents any crash inside DocumentsPage (or other pages) from unmounting
// the entire SPFx web part. Instead shows a friendly reset card.
class PageErrorBoundary extends React.Component<
  { pageName?: string; children: React.ReactNode },
  { hasError: boolean; errorMsg: string }
> {
  constructor(props: { pageName?: string; children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, errorMsg: '' };
  }
  static getDerivedStateFromError(error: unknown): { hasError: boolean; errorMsg: string } {
    const msg = error instanceof Error ? error.message : String(error);
    return { hasError: true, errorMsg: msg };
  }
  componentDidCatch(error: unknown, info: React.ErrorInfo): void {
    console.error('[VesselDMS] PageErrorBoundary caught render error:', error, info?.componentStack);
  }
  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div style={{
          margin: 24, padding: 32, borderRadius: 14, border: '1px solid #fecaca',
          background: '#fff5f5', color: '#991b1b', fontFamily: 'sans-serif',
        }}>
          <div style={{ fontSize: 22, marginBottom: 8 }}><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 22 }} /> Something went wrong</div>
          <div style={{ fontSize: 13, marginBottom: 4, fontWeight: 600 }}>
            {this.props.pageName || 'Page'} encountered an error and could not render.
          </div>
          <div style={{ fontSize: 12, color: '#b91c1c', marginBottom: 20, fontFamily: 'monospace', wordBreak: 'break-all' }}>
            {this.state.errorMsg}
          </div>
          <button
            type="button"
            onClick={() => this.setState({ hasError: false, errorMsg: '' })}
            style={{
              background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: 8,
              padding: '8px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer', boxShadow: clay.shadowButton,
            }}
          >
            Reset View
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

class DocumentsPageWrapper extends React.Component<{ host: any }> {
  public render(): React.ReactElement | null {
    return renderDocumentsPage(this.props.host);
  }
}

// ── State Interface ──────────────────────────────────────────────────────────

interface State {
  rows: FlatRow[];
  vessels: VesselRecord[];
  // Tenant-wide vessel names (every site, unscoped), fetched once purely so
  // DocumentsPage's search-clause classifier (matchesSearchTokens'
  // knownVesselNamesLower) can recognize "Bow Fighter + plan" as naming a
  // vessel even while browsing a SharePoint site whose `vessels` (the
  // Vessels-module table, scoped by vesselSiteFilter — see
  // _switchDocumentSite's comment on why it deliberately never refetches
  // that on a site switch) doesn't happen to include that vessel. Without
  // this, a clause naming a real vessel not in the currently-loaded
  // `vessels` list silently fell through to matchesSearchTokens' "unknown
  // clause" OR fallback, which matched the OTHER clause's words (e.g.
  // "plan") anywhere at all — pulling in unrelated vessels' files instead
  // of narrowing to the named one. Never used for anything else (no UI
  // renders from this list), so scoping it tenant-wide is safe.
  allVesselNamesForSearch: string[];
  loading: boolean;
  error: string | null;
  reloadKey: number;
  textFilter: string;
  // Documents page search box results dropdown (SearchResultsDropdown) —
  // open/closed state and the keyboard-highlighted row index (-1 = none).
  searchDropdownOpen: boolean;
  searchDropdownActiveIndex: number;
  vesselFilter: string;
  groupFilter: string;
  catFilter: string;
  attachmentFilter: 'all' | 'attached' | 'not_attached';
  sort: 'default' | 'name_asc' | 'name_desc';
  uploadingGroupKey: string | null;
  uploadInfo: string | null;
  uploadError: string | null;
  selectedFileIds: Set<string>;
  page: number;

  // Modal
  modal: ModalMode;
  vesselActionPicker: 'edit' | 'delete' | null;
  selectedVessel: VesselRecord | null;
  deleteVesselIds: Set<string>;
  deleteVesselProgress: Record<string, { name: string; status: 'waiting' | 'deleting' | 'success' | 'pending' | 'failed'; message?: string }>;
  /** Optional "Reason for deletion" text captured in the delete-vessel modal. */
  deleteVesselReason: string;
  form: FormState;
  modalBusy: boolean;
  modalMsg: string | null;
  modalError: string | null;
  deleteAutoCloseSeconds: number | null;
  formFieldErrors: Record<string, string>;

  // Active Main Navigation View
  view: AppView;
  recycleBin: DeletedNode[];
  archiveList: DeletedNode[];
  // Drive item IDs currently archived (GET /api/archive/ids) — used to hide
  // archived files from the normal Documents/Vessels views (they're only
  // "retrievable" via the Archive page's Restore action) without having to
  // re-derive that from archiveList, which only loads on the Archive page.
  archivedFileIds: Set<string>;
  archiveSearch: string;
  panelLoading: boolean;
  vesselsSearch: string;
  vesselStatusFilter: string;
  vesselTypeFilter: string;
  vesselSiteFilter: string;

  // Module Specific Data
  documentsList: DocumentItem[];
  documentSites: Array<{ site_key: string; sp_site_name: string; site_id: string; drive_id: string; web_url?: string; default_library_name?: string; is_primary?: boolean }>;
  documentDepartmentAliases?: Record<string, string[]>;
  documentVesselAliases?: Record<string, string[]>;
  tenantDiscoveredSites?: Array<{ id: string; name: string; display_name: string; web_url?: string }>;
  activeDocumentSite: string | null;
  documentLiveFolders: Array<{
    id: string;
    name: string;
    path: string;
    parent_id: string;
    depth: number;
    is_folder?: boolean;
    size?: number;
    created_date_time?: string;
    last_modified_date_time?: string;
    web_url?: string;
  }>;
  documentLiveFoldersLoading: boolean;
  docViewMode: 'folder' | 'list';
  docScopeType: 'vessels' | 'common' | 'kaizen' | 'sites' | 'shared_docs' | 'documents';
  docMainFolder: 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank' | 'Shared Documents' | 'Documents' | 'SharePoint Sites' | null;
  showAllVesselsInFolderView: boolean;
  docListPage: number;
  docListSort: 'name_az' | 'newest' | 'default';
  docGroupFilter: string;
  docCategoryFilter: string;
  docGroupLevelFilter: string;
  docLeafCategoryFilter: string;
  docSubCategoryFilter: string;
  // Live per-site folders that are neither a known main folder/category nor a
  // recognised vessel (e.g. "Report", "ACRA CHARGE", "Share with Mr Akase" —
  // real folders that sit alongside vessel folders in SharePoint but aren't
  // part of the app's department template). Selecting one filters Documents
  // down to that folder's own rows in both List view and Folder view.
  docSubfolderOtherFilter: string;
  // "Compare" mode: when a selected vessel has a folder under 2+ main
  // folders (Technical & Crewing, Commercial & Chartering, ...), this shows
  // one browsable box per main folder side by side instead of Folder view's
  // single breadcrumb-driven location. Opt-in via a toolbar toggle; Folder
  // view / List view are unaffected when it's off (the default).
  docCompareMode: boolean;
  // Per-main-folder browsing state for Compare mode: each box drills through
  // its own vessel-folder subtree independently of the others and of the
  // main folderPathStack. Keyed by main folder name; value is the chain of
  // sub-folder names navigated into within that box, relative to the
  // vessel's own folder under that main folder (empty = showing the vessel
  // folder's direct children).
  compareBoxSubPaths: Record<string, string[]>;
  documentVesselCount: number;
  documentVesselsLoadingMore: boolean;
  documentFilesLoading: boolean;
  vesselLoadingName: string | null;
  docUploadRowKey: string | null;
  docUploadBusy: boolean;
  docUploadMsg: string | null;
  // Phase 4 — full vessel + folder summary Excel export (independent of Documents filters)
  vesselsExcelExportBusy: boolean;
  folderPathStack: { id: string; name: string }[];
  uploadedFilesByFolder: Record<string, { name: string; size: string; date: string; pending?: boolean; uploading?: boolean; id?: string; uploadedAt?: number }[]>;
  selectedDocPreview: DocPreviewItem | null;
  templatesList: TemplateItem[];
  approvalsList: ApprovalItem[];
  approvalTab: 'Pending' | 'Approved' | 'Rejected';
  reportsSelectedVessel: string;
  ocrStagingCount: number;
  // Top-header alert bell — new folder/vessel creation alerts (replaces bottom-of-module notifications)
  alertsList: AlertItem[];
  // True once GET /api/alerts/all has resolved at least once. The deletion
  // toast layer (AppLayout.tsx) uses this — not "alertsList is non-empty" —
  // to tell "no alerts yet" (still loading) apart from "no deletions" so it
  // never mistakes old, already-existing deletion alerts for fresh ones.
  alertsLoaded: boolean;
  alertFilter: 'all' | 'unread';
  alertCategory: 'all' | 'dms' | 'crud' | 'email';
  selectedAlertId: string | null;
  alertOpen: boolean;
  usersList: UserItem[];
  userSearch: string;
  userRoleFilter: string;
  // Which user row's site-permission editor is expanded (email, or null).
  usersExpandedEmail: string | null;
  // In-progress edits for the expanded row's permissions, keyed by site_key,
  // before "Save" commits them via _updateUserSitePermissions.
  usersPermissionsDraft: UserSitePermissionItem[];
  usersPermissionsBusy: boolean;

  // Settings module state
  settingsTab: 'General' | 'Vessel Site Provisioning' | 'Document Settings' | 'Notification Settings' | 'Permission Settings' | 'Site Management' | 'Vessel Settings' | 'Tag Configuration' | 'Module Management' | 'Filter Search Management' | 'Color Management' | 'Audit Logs' | 'Settings Management';
  // Module ids (Sidebar.tsx navItems/auxLinks `id`) currently hidden
  // app-wide via Settings → Module Management. Loaded once in
  // componentDidMount (_loadModuleSettings) and refreshed immediately after
  // a save from that section. Sidebar filters its nav with this; _goToView
  // and the view switch below refuse to land on a hidden module.
  hiddenModules: string[];
  // Settings-page tab labels (SettingsPage.tsx settingsTab values) currently
  // hidden app-wide via Settings → Settings Management. Loaded once in
  // componentDidMount (_loadSettingsTabSettings) and refreshed immediately
  // after a save from that section. SettingsPage.tsx filters its own left
  // nav with this and redirects off a tab that just got hidden.
  hiddenSettingsTabs: string[];
  // How the Documents page's filter toolbar is presented, app-wide, via
  // Settings → Filter Search Management (filter_settings_api.py). Loaded
  // once in componentDidMount (_loadFilterSettings) and refreshed
  // immediately after a save from that section. DocumentsPage.tsx reads
  // this to switch between the inline dropdown row ('dropdown') and the
  // slide-out Filters panel ('panel').
  filterUiMode: 'dropdown' | 'panel';
  // Whether the slide-out Filters panel (filterUiMode === 'panel') is
  // currently open on the Documents page.
  filterPanelOpen: boolean;
  // Background/text/design(accent)/hover colors, per light+night mode, set
  // app-wide via Settings → Color Management (color_settings_api.py).
  // Loaded once in componentDidMount (_loadColorSettings) and applied via
  // applyColorTheme(); refreshed immediately after a save from that
  // section. AppLayout.tsx also reads this to rebuild Fluent UI's theme.
  colorTheme: ClayColorTheme;
  settingsForm: {
    siteTitle: string;
    siteDescription: string;
    dateFormat: string;
    timeZone: string;
  };
  settingsSavedMsg: boolean;

  // AI Bento Email Status & Notification module
  bentoConfigRecipient: string;
  bentoLogs: BentoEmailLog[];
  bentoStatusFilter: string;
  bentoSearch: string;
  bentoComposeOpen: boolean;
  bentoComposeForm: {
    vessel_name: string;
    datasource_tag: string;
    subject_text: string;
    body: string;
    file: File | null;
    recipient: string;
    existing_attachment: string;
  };
  bentoComposeBusy: boolean;
  bentoComposeMsg: string | null;
  bentoComposeErr: string | null;
  bentoDetailLog: BentoEmailLog | null;
  bentoClearAllBusy: boolean;
  bentoClearAllErr: string | null;

  // Approved files cache per vessel (keyed by lowercase vessel name)
  bentoApprovedFiles: Record<string, string[]>;
  bentoApprovedFileIds: Record<string, string>;
  // Bento Upload Widget State
  bentoUploadFile: File | null;
  bentoUploadVessel: string;
  bentoUploadTag: string;
  bentoUploadBusy: boolean;
  bentoUploadMsg: string | null;
  bentoUploadErr: string | null;

  // SharePoint folder creation state
  folderCreationBusy: boolean;
  folderCreationResults: FolderResult[] | null;
  folderCreationError: string | null;
  folderProvisioningVesselId: string | null;
  folderCreationFeed: FolderResult[];
  provisionedVesselIds: Set<string>;

  // Vessel auto-discovery sync (SharePoint root folders -> DB/Term Store).
  // The manual "Sync Vessels from SharePoint" button (and vesselSyncRunning/
  // vesselSyncSummary, which only it ever set) was removed — the backend
  // already runs this in the background on every vessel-list fetch (see
  // _maybeAutoSyncVesselsFromSharePoint). vesselSyncError stays: it's also
  // used by _confirmDiscoveredVessel for its own failures.
  vesselSyncError: string | null;
  // Draft edits (IMO / Hull No.) for "Found in SharePoint" cards, keyed by a
  // stable key (site_key + folder path) since these vessels have no DB id yet
  discoveredVesselDrafts: Record<string, { imo: string; hull_number: string }>;
  // Key of the discovered-vessel card currently being confirmed (POST in flight)
  confirmingVesselKey: string | null;
  // Per-card error for a failed "Confirm Vessel" click, keyed by
  // _discoveredVesselKey. Shown inline under that card's button — the
  // page-level vesselSyncError banner at the top of the page is easy to
  // miss once the user has scrolled down into the vessel grid, which made
  // a failed confirm look like it silently did nothing ("no response").
  discoveredVesselConfirmError: Record<string, string>;

  // In-progress edits for a regular ("dms"-source) vessel card's IMO / Hull
  // No. / Shipyard / Type fields, offered only when that field is currently
  // empty (see renderVesselsPage). Keyed by vessel.id, unlike
  // discoveredVesselDrafts which is deferred-save; these commit immediately
  // via _saveVesselField (PATCH /api/vessels/{id}) on blur/select.
  vesselFieldDrafts: Record<string, { imo?: string; hull_number?: string; shipyard?: string; vessel_type?: string }>;
  // Save-in-flight / last-error state per "<vesselId>:<field>" key.
  vesselFieldSaving: Record<string, boolean>;
  vesselFieldError: Record<string, string>;
  // A single field's edit awaiting the user's confirmation before it is
  // PATCHed and locked. Set by _requestSaveVesselField (blur/select), cleared
  // by confirming (which then calls _saveVesselField) or cancelling (which
  // reverts the draft). Only one confirmation can be open at a time.
  vesselFieldConfirm: {
    vesselId: string;
    vesselName: string;
    field: 'imo' | 'hull_number' | 'shipyard' | 'vessel_type';
    fieldLabel: string;
    value: string;
  } | null;

  // Delta sync — flat id→node map representing the live SPO folder tree
  spoFolderMap: Map<string, SpoFolderNode>;
  lastDeltaSync: Date | null;

  sessionExpired: boolean;
  authPage: AuthPageMode | null;
  sessionReady: boolean;  // true once first valid session_id prop is received

  // Toast shown when a vessel is auto-moved to recycle bin via SPO deletion
  spoVesselDeletedToast: {
    vesselName: string;
    id: string;
    siteName?: string | null;
    originalPath?: string | null;
    deletedByName?: string | null;
    deletedByEmail?: string | null;
  } | null;
  // Toast shown when a document file/folder is moved to the SPO recycle bin.
  // `items` carries per-item attribution (path / site / who deleted it) so the
  // popup and the Alerts detail panel can both show "who / what / where".
  spoDocumentDeletedToast: {
    itemType: 'file' | 'folder';
    items: Array<{
      id: string;
      name: string;
      path?: string | null;
      siteName?: string | null;
      deletedByName?: string | null;
      deletedByEmail?: string | null;
    }>;
  } | null;
  // IDs of vessels soft-deleted via SPO — used to filter them out of re-fetched vessel lists
  spoDeletedVesselIds: Set<string>;

  // Folder placement anomalies
  folderAnomalies: FolderAnomalyItem[];

  // Folders that users have confirmed as "Normal Folders" (not vessels)
  normalFolders: NormalFolderRecord[];

  // SPO item classify dialog (vessel vs normal folder)
  spoClassifyDialog: {
    anomaly: FolderAnomalyItem;
    provisioning: boolean;
    done: boolean;
    doneNormal?: boolean;   // success screen for "normal folder" classification
    alreadyExisted?: boolean;
    error: string | null;
  } | null;

  // Vessel card Provision dialog (standalone provision for existing vessel)
  spoProvisionDialog: {
    vessel: import('./types/rows').VesselRecord;
    provisioning: boolean;
    done: boolean;
    error: string | null;
  } | null;

  // Dismiss-with-confirm dialog — item being confirmed for recycle bin move
  spoAnomalyDismissConfirm: FolderAnomalyItem | null;

  // File alert dialog — file uploaded directly to SPO outside DMS structure
  spoFileAlertDialog: import('./types/ui').SpoFileAlertDialog | null;

  // Sidebar collapse/expand state
  sidebarCollapsed: boolean;
  themeMode: 'light' | 'night';
  fullScreenWorkspace: boolean;
  /** Top-bar "Documents" button: shows the traditional SharePoint library inside the app. */
  classicSiteOpen: boolean;
  classicSiteUrl: string;
  windowWidth: number;

  // Folder navigation history (back/forward)
  folderNavHistory: Array<{ folderPathStack: { id: string; name: string }[]; docMainFolder: State['docMainFolder'] }>;
  folderNavIndex: number;

  // File delete dialog
  fileDeleteDialog: {
    files: Array<{ id: string; name: string; folderId: string; folderPath: string; vesselName?: string; siteId?: string; driveId?: string }>;
    selected: Set<string>;
    busy: boolean;
    error: string | null;
    /** "Reason for deletion (optional)" text captured before confirming. */
    reason?: string;
  } | null;

  // Archive picker popup — opened from the "Archive" toolbar button when
  // browsing the plain SharePoint folder tree (Sites / Shared Documents /
  // Documents), which has no per-file checkbox selection of its own. Lets
  // the user check folders (recursively) and/or individual files under the
  // folder they're currently viewing, then archives everything checked.
  archivePickerDialog: ArchivePickerDialogState | null;

  // Mirrors archivePickerDialog above, but for the "New Folder" toolbar
  // button (top, next to Archive) — opens a popup to create a folder at
  // (or under an edited variant of) whatever path the user is currently
  // viewing, with an option to redirect into the Add Vessel form instead.
  createFolderDialog: CreateFolderDialogState | null;

  // List view per-file checkbox selection (keyed by file id)
  listViewSelectedFiles: Set<string>;

  // Folder view per-file checkbox selection (keyed by file id)
  folderViewSelectedFiles: Set<string>;

  // Upload success popup
  uploadSuccessPopup: {
    fileName: string;
    destinationPath: string;
    webUrl: string;
    isPending: boolean;
    secondsLeft: number;
    unidentifiedFiles?: string[];
  } | null;

  // Post-upload navigation prompt when OCR cannot identify vessel names
  uploadNavigationPrompt: {
    unidentifiedCount: number;
    unresolvedFiles: string[];
    queueTab: 'staged' | 'unstaged';
    dontAskAgain: boolean;
  } | null;

  // User preference: skip post-upload OCR navigation prompt and stay on current page
  suppressUploadNavigationPrompt: boolean;

  // Bulk Upload Modal Dialog
  bulkUploadDialog: {
    files: BulkUploadFile[];
    folderId: string;
    subFolderPath: string;
    vesselName: string;
    currentFolderNode: { id: string; name: string } | null;
    targetSiteId?: string;
    targetDriveId?: string;
  } | null;

  // Add Folder Dialog (create-subfolder under any vessel-owned or
  // month-driven folder — Phase 3 "Add Folder" flow)
  addFolderDialog: {
    folderId: string;
    folderLabel: string;
    vesselName: string;
    name: string;
    busy: boolean;
    error: string | null;
    /** When checked, submitting redirects to the full Add Vessel form
     *  (pre-filled with `name`) instead of creating a plain subfolder. */
    asVessel: boolean;
  } | null;

  // Folder Delete Dialog (moves uploaded folder to Recycle Bin)
  folderDeleteDialog: {
    currentFolderId: string;
    currentFolderName: string;
    currentFolderPath: string;
    vesselName: string;
    mainFolder: string;
    availableFolders: Array<{ id: string; name: string; path: string; isCurrent?: boolean }>;
    selectedFolderId: string;
    busy: boolean;
    error: string | null;
    /** "Reason for deletion (optional)" text captured before confirming. */
    reason?: string;
  } | null;

  // OCR pending file from List/Folder View (triggers OCR tab in Templates)
  ocrPendingItemId: string | null;
  ocrPendingFilename: string | null;
  ocrQueueTabHint: 'staged' | 'unstaged' | null;
  ocrUnidentifiedFiles: string[];

  // Vessel Suggestions wizard dialog
  vesselSuggestionDialog: VesselSuggestionDialog | null;
  // Unresolved new-vessel suggestions retained from latest scans/uploads
  pendingVesselSuggestions: VesselSuggestion[];
  // Suggestion keys marked as false positive and ignored by user
  ignoredVesselSuggestionKeys: Set<string>;

  // Sidebar vessels drilldown dropdown state
  vesselsNavExpanded: boolean;

  // Real-time cached Dashboard/Home counts from backend
  dashboardStats: DashboardStats | null;
  // SharePoint site the Dashboard's stat cards are scoped to ('all' = every
  // configured site, merged). Independent of vesselSiteFilter/activeDocumentSite
  // so switching it doesn't disturb the Vessels/Documents site selections.
  dashboardSiteFilter: string;

  // Persistent progress for long-running site OCR scans, including navigation away from Sites.
  scanProgress: {
    status: 'idle' | 'running' | 'completed' | 'failed';
    completed: number;
    total: number;
    title: string;
    recentFiles: string[];
    error?: string;
  };
}


const BLANK_FORM: FormState = { name: '', imo: '', shipyard: '', hull_number: '', vessel_type: '', target_site_ids: [] };

const PAGE_SIZE = 50;

// ── Component Definition ────────────────────────────────────────────────────

export default class VesselEmail extends React.Component<IVesselEmailProps, State> {

  public _abort: AbortController | null = null;
  public _filesLoadedForVessels: Set<string> = new Set();
  // Debounce timer + already-searched-terms cache for the Documents search
  // box's fleet-wide backend lookup (see _scheduleGlobalSearch /
  // _triggerGlobalSearch below) — lets a search term discover matches in
  // vessels the user hasn't opened yet, instead of only filtering rows
  // already loaded into state.
  public _globalSearchTimer: ReturnType<typeof setTimeout> | null = null;
  public _globalSearchSeenTerms: Set<string> = new Set();
  public _appUploadedFileIds: Set<string> = new Set();
  public _syncScheduler: SyncScheduler | null = null;
  private _folderRefreshSeq = 0;
  private _latestFolderRefreshTokenByKey: Map<string, number> = new Map();
  private _documentSitesLoading = false;
  private _documentLiveTreeLoading: Set<string> = new Set();
  private _documentLiveTreeLoaded: Set<string> = new Set();
  private _documentLiveTreeAbortControllers: Map<string, AbortController> = new Map();
  private _liveSharePointMerges: Set<string> = new Set();
  // Monotonic request counter for _loadDashboardStats. Guards against a
  // slow in-flight request (e.g. a full live scan for a previous site
  // selection) resolving AFTER a newer one and clobbering it. Sequencing
  // the requests themselves — rather than comparing last_refreshed_epoch,
  // which is a per-site server cache timestamp and isn't comparable across
  // different sites — is what lets a dashboard site switch apply immediately
  // instead of waiting for the newly selected site's cache epoch to catch up.
  private _dashboardStatsSeq = 0;
  private _dashboardStatsAppliedSeq = 0;
  public _isLoadingData = false;
  public _isUnmounted = false;
  // Throttle state for the background auto-sync-from-SharePoint (see
  // _maybeAutoSyncVesselsFromSharePoint): last time (ms) an automatic sync
  // ran per site_key ('all' for the unfiltered view), so _loadData — which
  // can fire many times a minute (delta reload, polling, user actions) —
  // doesn't re-trigger a full SharePoint scan on every call.
  private _lastAutoSyncAt: Record<string, number> = {};
  private static readonly AUTO_SYNC_MIN_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
  public _deltaReloadTimer: ReturnType<typeof setTimeout> | null = null;
    public _deltaFileRefreshTimer: ReturnType<typeof setTimeout> | null = null;   // ← add this line
  public _alertRefreshTimer: ReturnType<typeof setInterval> | null = null;
  public _deleteAutoCloseTimer: ReturnType<typeof setInterval> | null = null;
  // Auto-dismiss timers for the two SPO-deletion popups (Recycle Bin fix #5):
  // each popup dismisses itself after _TOAST_AUTO_DISMISS_MS unless the
  // pointer is over it, in which case the timer is cleared on hover and
  // restarted on mouse-leave.
  public _vesselToastAutoCloseTimer: ReturnType<typeof setTimeout> | null = null;
  public _documentToastAutoCloseTimer: ReturnType<typeof setTimeout> | null = null;
  private static readonly _TOAST_AUTO_DISMISS_MS = 7000;
  // Drive item IDs we've already raised a deletion popup for, across BOTH
  // the vessel-deletion and generic document-deletion paths — belt-and-
  // suspenders against the same SPO deletion event surfacing more than one
  // popup (e.g. a vessel's root folder being reported once as a vessel
  // deletion and again by a generic delta-sync/poll path). Kept as a plain
  // instance Set (not React state) since it's bookkeeping, not something
  // that should trigger a re-render or ever need to be reset by the user.
  public _shownDeletionToastIds: Set<string> = new Set();
  public _appDeletedItemIds: Set<string> = new Set();
  public _initialDocumentFolderRefreshDone = false;
  private readonly _uploadCacheVersion = 1;
  public _handleResize = (): void => {
    const nextWidth = window.innerWidth;
    this.setState(prev => ({
      windowWidth: nextWidth,
      sidebarCollapsed: isTabletOrBelow(nextWidth) ? true : prev.sidebarCollapsed,
    }));
  };

  // ── Folder structure constants ────────────────────────────────────────────
  // New primary structure: Documents/{MainFolder}/{VesselName}/...
  public readonly MAIN_FOLDER_NAMES: string[] = [
    'Technical & Crewing',
    'Commercial & Chartering',
    'Insurance',
  ];
  // Root prefix under which vessels/common reside (empty = Documents root)
  public readonly VESSEL_ROOT: string = '';

  // Tracks whether the user has already clicked into the Documents module
  // ('list' view) since this component was constructed (i.e. since the last
  // full page load / hard refresh). Always starts false on a fresh mount —
  // see _goToView's 'list' branch for why this matters.
  private _hasEnteredDocumentsSinceMount: boolean = false;

  // Set once _loadDocumentSites (componentDidMount) has finished resolving
  // the initial/restored site+drive+folder root against live Graph data —
  // i.e. state.folderPathStack/docScopeType are already trustworthy, not a
  // stale unverified leftover. _goToView's 'list' branch used to always
  // discard a live 'sites'/'shared_docs'/'documents' scope on the user's
  // very first "Documents" click after mount (isFirstEntrySinceMount),
  // assuming any such scope that early could only be an unverified
  // sessionStorage restore. But _loadDocumentSites can finish (and set a
  // freshly-verified 'sites' root, see its wantsInitialSiteRoot/live-drive
  // logic) before that first click happens, and the reset threw that good
  // data away — landing the user on the empty vessels-DB root ("0 main
  // folders & libraries") on hard refresh / app start until they manually
  // reselected a site (which re-resolves the drive via
  // _switchDocumentSite). This flag lets _goToView tell the two cases
  // apart instead of always assuming the worst on the first click.
  private _initialDocumentsRootReady: boolean = false;

  public constructor(props: IVesselEmailProps) {
    super(props);
    const ignoredSuggestionKeys = this._readIgnoredVesselSuggestionKeys();
    const restoredPendingSuggestions = this._readPendingVesselSuggestions().filter(
      suggestion => !ignoredSuggestionKeys.has(this._getSuggestionKey(suggestion))
    );
    const suppressUploadNavigationPrompt = this._readSuppressUploadNavigationPrompt();
    const savedThemeMode = typeof window !== 'undefined' ? window.localStorage.getItem('vesseldms.theme-mode') : null;
    const themeMode: 'light' | 'night' = savedThemeMode === 'night' ? 'night' : 'light';
    const persistedFolderNav = this._readPersistedFolderNav();
    this.state = {
      rows: [],
      vessels: [],
      allVesselNamesForSearch: [],
      loading: false, error: null, reloadKey: 0, sessionExpired: Boolean(props.sessionExpired),
      authPage: props.sessionExpired ? 'login' : null,
      textFilter: '', searchDropdownOpen: false, searchDropdownActiveIndex: -1,
      vesselFilter: 'all', groupFilter: 'all', catFilter: 'all', attachmentFilter: 'all',
      sort: 'default', uploadingGroupKey: null, uploadInfo: null, uploadError: null,
      selectedFileIds: new Set(), page: 0,
           modal: 'none', selectedVessel: null, deleteVesselIds: new Set(), deleteVesselProgress: {}, deleteVesselReason: '', form: { ...BLANK_FORM },
           vesselActionPicker: null,
      modalBusy: false, modalMsg: null, modalError: null, deleteAutoCloseSeconds: null, formFieldErrors: {},
      view: 'dashboard',
      recycleBin: [], archiveList: [], archivedFileIds: new Set<string>(), archiveSearch: '', panelLoading: false,
      vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all', vesselSiteFilter: 'all',

      documentsList: INITIAL_MOCK_DOCUMENTS,
      documentSites: [],
      documentDepartmentAliases: {},
      documentVesselAliases: {},
      tenantDiscoveredSites: [],
      activeDocumentSite: persistedFolderNav ? persistedFolderNav.activeDocumentSite : null,
      documentLiveFolders: [],
      documentLiveFoldersLoading: false,
      docViewMode: persistedFolderNav ? persistedFolderNav.docViewMode : 'folder',
      docScopeType: persistedFolderNav ? persistedFolderNav.docScopeType : 'vessels',
      docMainFolder: persistedFolderNav ? persistedFolderNav.docMainFolder : null,
      showAllVesselsInFolderView: false,
      docListPage: 0,
      docListSort: 'default',
      docGroupFilter: 'all',
      docCategoryFilter: persistedFolderNav ? persistedFolderNav.docCategoryFilter : 'all',
      docGroupLevelFilter: 'all',
      docLeafCategoryFilter: 'all',
      docSubCategoryFilter: 'all',
      docSubfolderOtherFilter: persistedFolderNav ? persistedFolderNav.docSubfolderOtherFilter : 'all',
      docCompareMode: false,
      compareBoxSubPaths: {},
      documentVesselCount: 4,
      documentVesselsLoadingMore: false,
      documentFilesLoading: false,
      vesselLoadingName: null,
      docUploadRowKey: null,
      docUploadBusy: false,
      vesselsExcelExportBusy: false,
      docUploadMsg: null,
      folderPathStack: persistedFolderNav ? persistedFolderNav.folderPathStack : [],
      uploadedFilesByFolder: {},
      selectedDocPreview: null,
      templatesList: INITIAL_MOCK_TEMPLATES,
      approvalsList: INITIAL_MOCK_APPROVALS,
      approvalTab: 'Pending',
      reportsSelectedVessel: 'All Vessels',
      ocrStagingCount: 0,
      alertsList: [],
      alertsLoaded: false,
      alertFilter: 'all',
      alertCategory: 'all',
      selectedAlertId: null,
      alertOpen: false,
      usersList: [],
      userSearch: '', userRoleFilter: 'all',
      usersExpandedEmail: null,
      usersPermissionsDraft: [],
      usersPermissionsBusy: false,

      settingsTab: 'Site Management',
      hiddenModules: [],
      hiddenSettingsTabs: [],
      colorTheme: DEFAULT_CLAY_COLORS,
      filterUiMode: 'dropdown',
      filterPanelOpen: false,
      settingsForm: {
        siteTitle: 'Vessel Documents Management',
        siteDescription: 'Manage and track all vessel related documents efficiently.',
        dateFormat: 'MM/DD/YYYY',
        timeZone: '(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi',
      },
      settingsSavedMsg: false,

      bentoConfigRecipient: '',
      bentoLogs: [], bentoStatusFilter: 'all', bentoSearch: '',
      bentoComposeOpen: false,
      bentoComposeForm: { vessel_name: '', datasource_tag: 'mail', subject_text: '', body: '', file: null, existing_attachment: '', recipient: '' },
      bentoComposeBusy: false, bentoComposeMsg: null, bentoComposeErr: null,
      bentoDetailLog: null,
      bentoClearAllBusy: false, bentoClearAllErr: null,
      bentoUploadFile: null, bentoUploadVessel: '', bentoUploadTag: 'mail',
      bentoUploadBusy: false, bentoUploadMsg: null, bentoUploadErr: null,
      bentoApprovedFiles: {},
      bentoApprovedFileIds: {},

      folderCreationBusy: false,
      folderCreationResults: null,
      folderCreationError: null,
      folderProvisioningVesselId: null,
      folderCreationFeed: [],
      provisionedVesselIds: new Set<string>(),
      vesselSyncError: null,
      discoveredVesselDrafts: {},
      confirmingVesselKey: null,
      discoveredVesselConfirmError: {},
      vesselFieldDrafts: {},
      vesselFieldSaving: {},
      vesselFieldError: {},
      vesselFieldConfirm: null,
      spoFolderMap: new Map(),
      lastDeltaSync: null,
      sessionReady: false,

      folderAnomalies: [],
      normalFolders: [],
      spoClassifyDialog: null,
      spoProvisionDialog: null,
      spoAnomalyDismissConfirm: null,
      spoFileAlertDialog: null,
      spoVesselDeletedToast: null,
      spoDocumentDeletedToast: null,
      spoDeletedVesselIds: new Set<string>(),
      sidebarCollapsed: false,
      themeMode,
      fullScreenWorkspace: false,
      classicSiteOpen: false,
      classicSiteUrl: '',
      windowWidth: typeof window !== 'undefined' ? window.innerWidth : 1200,
      folderNavHistory: persistedFolderNav
        ? [{ folderPathStack: persistedFolderNav.folderPathStack, docMainFolder: persistedFolderNav.docMainFolder }]
        : [{ folderPathStack: [], docMainFolder: null }],
      folderNavIndex: 0,
      uploadSuccessPopup: null,
      uploadNavigationPrompt: null,
      bulkUploadDialog: null,
      fileDeleteDialog: null,
      archivePickerDialog: null,
      createFolderDialog: null,
      addFolderDialog: null,
      folderDeleteDialog: null,
      listViewSelectedFiles: new Set<string>(),
      folderViewSelectedFiles: new Set<string>(),
      ocrPendingItemId: null,
      ocrPendingFilename: null,
      ocrQueueTabHint: null,
      ocrUnidentifiedFiles: [],
      vesselSuggestionDialog: null,
      pendingVesselSuggestions: restoredPendingSuggestions,
      ignoredVesselSuggestionKeys: ignoredSuggestionKeys,
      suppressUploadNavigationPrompt,
      vesselsNavExpanded: true,
      dashboardStats: null,
      dashboardSiteFilter: 'all',
      scanProgress: {
        status: 'idle', completed: 0, total: 0, title: '', recentFiles: [],
      },
    };
  }

  public _toggleThemeMode = (): void => {
    this.setState(prev => {
      const themeMode: 'light' | 'night' = prev.themeMode === 'light' ? 'night' : 'light';
      if (typeof window !== 'undefined') {
        window.localStorage.setItem('vesseldms.theme-mode', themeMode);
      }
      return { themeMode };
    });
  };

  public _toggleFullScreenWorkspace = (): void => {
    this.setState(prev => ({ fullScreenWorkspace: !prev.fullScreenWorkspace }));
  };

  /** Show the site's traditional "Documents" library inside the app (top-bar
   * "Documents" button). The URL is the drive's own webUrl from Graph — the
   * library's URL segment is usually "Shared Documents", not "Documents" —
   * with the site URL as a fallback. The Vessel DMS content stays mounted
   * underneath, so _closeClassicSite returns to exactly where the user was. */
  public _openClassicSite = async (): Promise<void> => {
    // Follow the site the user is looking at: the Dashboard / Vessels
    // "SharePoint site" filters, or the Documents site elsewhere. "All
    // sites" (or nothing chosen) falls back to the site the web part is on.
    const { view, dashboardSiteFilter, vesselSiteFilter, activeDocumentSite, documentSites } = this.state;
    const chosen = view === 'dashboard' ? dashboardSiteFilter
      : view === 'vessels' ? vesselSiteFilter
      : activeDocumentSite;
    const key = chosen && chosen !== 'all' ? chosen.toLowerCase() : '';
    const site = key ? (documentSites || []).find(s => String(s.site_key).toLowerCase() === key) : undefined;

    let url = '';
    try {
      const { graphClient } = this.props;
      const driveId = site?.drive_id || this.props.driveId;
      if (graphClient && driveId) {
        const drive = await graphClient.api(`/drives/${driveId}`).select('webUrl').get();
        url = drive?.webUrl || '';
      }
    } catch (e) {
      console.warn('[VesselDMS] could not resolve the Documents library URL:', e);
    }
    if (!url) url = `${(site?.web_url || this.props.siteUrl || '').replace(/\/$/, '')}/Shared%20Documents`;
    this.setState({ classicSiteOpen: true, classicSiteUrl: url });
  };

  public _closeClassicSite = (): void => {
    this.setState({ classicSiteOpen: false });
  };

  private _uploadCacheStorageKey(): string {
    const scope = (this.props.siteId || this.props.siteUrl || 'default').toString().toLowerCase();
    return `vesseldms.upload-cache:${scope}:v${this._uploadCacheVersion}`;
  }

  private _pendingVesselSuggestionsStorageKey(): string {
    const scope = (this.props.siteId || this.props.siteUrl || 'default').toString().toLowerCase();
    return `vesseldms.pending-vessel-suggestions:${scope}:v1`;
  }

  private _vesselSuggestionPrefsStorageKey(): string {
    const scope = (this.props.siteId || this.props.siteUrl || 'default').toString().toLowerCase();
    return `vesseldms.vessel-suggestion-prefs:${scope}:v1`;
  }

  private _readIgnoredVesselSuggestionKeys(): Set<string> {
    if (typeof window === 'undefined') return new Set<string>();
    try {
      const raw = window.localStorage.getItem(this._vesselSuggestionPrefsStorageKey());
      if (!raw) return new Set<string>();
      const parsed = JSON.parse(raw) as { ignoredKeys?: string[] };
      const keys = Array.isArray(parsed?.ignoredKeys) ? parsed.ignoredKeys.filter(Boolean) : [];
      return new Set<string>(keys);
    } catch {
      return new Set<string>();
    }
  }

  private _persistIgnoredVesselSuggestionKeys(keys: Set<string>): void {
    if (typeof window === 'undefined') return;
    try {
      const raw = window.localStorage.getItem(this._vesselSuggestionPrefsStorageKey());
      const parsed = raw ? JSON.parse(raw) as { suppressUploadNavigationPrompt?: boolean } : {};
      window.localStorage.setItem(this._vesselSuggestionPrefsStorageKey(), JSON.stringify({
        ...parsed,
        ignoredKeys: Array.from(keys.values()),
      }));
    } catch {
      // best-effort only
    }
  }

  private _readSuppressUploadNavigationPrompt(): boolean {
    if (typeof window === 'undefined') return true;
    try {
      const raw = window.localStorage.getItem(this._vesselSuggestionPrefsStorageKey());
      if (!raw) return true;
      const parsed = JSON.parse(raw) as { suppressUploadNavigationPrompt?: boolean };
      if (parsed && parsed.suppressUploadNavigationPrompt !== undefined) {
        return Boolean(parsed.suppressUploadNavigationPrompt);
      }
      return true;
    } catch {
      return true;
    }
  }

  private _persistSuppressUploadNavigationPrompt(suppress: boolean): void {
    if (typeof window === 'undefined') return;
    try {
      const raw = window.localStorage.getItem(this._vesselSuggestionPrefsStorageKey());
      const parsed = raw ? JSON.parse(raw) as { ignoredKeys?: string[] } : {};
      window.localStorage.setItem(this._vesselSuggestionPrefsStorageKey(), JSON.stringify({
        ...parsed,
        suppressUploadNavigationPrompt: Boolean(suppress),
      }));
    } catch {
      // best-effort only
    }
  }

  public _setSuppressUploadNavigationPrompt(suppress: boolean): void {
    const next = Boolean(suppress);
    this._persistSuppressUploadNavigationPrompt(next);
    this.setState(prev => ({
      suppressUploadNavigationPrompt: next,
      uploadNavigationPrompt: next ? null : prev.uploadNavigationPrompt,
    }));
  }

  // ── Folder navigation persistence (survive a page refresh) ─────────────────
  // sessionStorage (not localStorage): the selected folder/sub-folder should
  // come back after a refresh in the same tab, but shouldn't resurrect a
  // months-old navigation the next time the workbench page is opened fresh.
  private _folderNavStorageKey(): string {
    const scope = (this.props.siteId || this.props.siteUrl || 'default').toString().toLowerCase();
    return `vesseldms.folder-nav:${scope}:v1`;
  }

  private _readPersistedFolderNav(): {
    folderPathStack: { id: string; name: string }[];
    docMainFolder: State['docMainFolder'];
    docScopeType: State['docScopeType'];
    docViewMode: State['docViewMode'];
    docCategoryFilter: string;
    docSubfolderOtherFilter: string;
    activeDocumentSite: string | null;
  } | null {
    if (typeof window === 'undefined') return null;
    try {
      const raw = window.sessionStorage.getItem(this._folderNavStorageKey());
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      const stack = Array.isArray(parsed?.folderPathStack) ? parsed.folderPathStack : [];
      const validStack = stack.filter((n: any) => n && typeof n.id === 'string' && typeof n.name === 'string');
      if (validStack.length === 0) return null;
      const validMainFolders: State['docMainFolder'][] = [
        'Technical & Crewing', 'Commercial & Chartering', 'Insurance', 'Kaizen - Knowledge Bank',
        'Knowledge Bank', 'Shared Documents', 'Documents', 'SharePoint Sites',
      ];
      const validScopeTypes: State['docScopeType'][] = ['vessels', 'common', 'kaizen', 'sites', 'shared_docs', 'documents'];
      const docMainFolder = validMainFolders.indexOf(parsed?.docMainFolder) !== -1 ? parsed.docMainFolder as State['docMainFolder'] : null;
      const docScopeType = validScopeTypes.indexOf(parsed?.docScopeType) !== -1 ? parsed.docScopeType as State['docScopeType'] : 'vessels';
      return {
        folderPathStack: validStack,
        docMainFolder,
        docScopeType,
        docViewMode: parsed.docViewMode === 'list' ? 'list' : 'folder',
        docCategoryFilter: typeof parsed.docCategoryFilter === 'string' ? parsed.docCategoryFilter : 'all',
        docSubfolderOtherFilter: typeof parsed.docSubfolderOtherFilter === 'string' ? parsed.docSubfolderOtherFilter : 'all',
        activeDocumentSite: typeof parsed.activeDocumentSite === 'string' ? parsed.activeDocumentSite : null,
      };
    } catch {
      return null;
    }
  }

  private _persistFolderNav(): void {
    if (typeof window === 'undefined') return;
    try {
      const { folderPathStack, docMainFolder, docScopeType, docViewMode, docCategoryFilter, docSubfolderOtherFilter, activeDocumentSite } = this.state;
      if (!folderPathStack || folderPathStack.length === 0) {
        window.sessionStorage.removeItem(this._folderNavStorageKey());
        return;
      }
      window.sessionStorage.setItem(this._folderNavStorageKey(), JSON.stringify({
        folderPathStack, docMainFolder, docScopeType, docViewMode, docCategoryFilter, docSubfolderOtherFilter, activeDocumentSite,
      }));
    } catch {
      // best-effort only
    }
  }

  private _readPendingVesselSuggestions(): VesselSuggestion[] {
    if (typeof window === 'undefined') return [];
    try {
      const raw = window.localStorage.getItem(this._pendingVesselSuggestionsStorageKey());
      if (!raw) return [];
      const parsed = JSON.parse(raw) as { pending?: VesselSuggestion[] };
      const list = Array.isArray(parsed?.pending) ? parsed.pending : [];
      return list
        .filter(item => item && typeof item.vesselName === 'string')
        .map(item => ({
          ...item,
          matchedExisting: null,
          sourceFiles: Array.isArray(item.sourceFiles) ? item.sourceFiles : [],
          lastDetectedAt: Number(item.lastDetectedAt) || Date.now(),
        }))
        .sort((a, b) => (b.lastDetectedAt || 0) - (a.lastDetectedAt || 0));
    } catch {
      return [];
    }
  }

  private _persistPendingVesselSuggestions(pending: VesselSuggestion[]): void {
    if (typeof window === 'undefined') return;
    try {
      const safe = (pending || []).map(item => ({
        vesselName: item.vesselName || '',
        imoNumber: item.imoNumber || '',
        shipyard: item.shipyard || '',
        hullNumber: item.hullNumber || '',
        vesselType: item.vesselType || '',
        matchedExisting: null,
        confidence: item.confidence || 0,
        sourceFiles: Array.isArray(item.sourceFiles) ? item.sourceFiles : [],
        ocrVerified: !!item.ocrVerified,
        ocrTextPreview: item.ocrTextPreview || '',
        ocrConfidence: item.ocrConfidence || 0,
        matchedKeywords: Array.isArray(item.matchedKeywords) ? item.matchedKeywords : [],
        suggestedTags: item.suggestedTags || {},
        lastDetectedAt: Number(item.lastDetectedAt) || Date.now(),
      } as VesselSuggestion));
      window.localStorage.setItem(this._pendingVesselSuggestionsStorageKey(), JSON.stringify({ pending: safe }));
    } catch {
      // best-effort only
    }
  }

  private _readUploadCache(): { rows: FlatRow[]; uploadedFilesByFolder: State['uploadedFilesByFolder'] } | null {
    if (typeof window === 'undefined') return null;
    try {
      const raw = window.localStorage.getItem(this._uploadCacheStorageKey());
      if (!raw) return null;
      const parsed = JSON.parse(raw) as {
        version?: number;
        savedAt?: number;
        rows?: FlatRow[];
        uploadedFilesByFolder?: State['uploadedFilesByFolder'];
      };
      if (!parsed || parsed.version !== this._uploadCacheVersion) return null;
      const maxAgeMs = 14 * 24 * 60 * 60 * 1000;
      if (parsed.savedAt && Date.now() - parsed.savedAt > maxAgeMs) {
        window.localStorage.removeItem(this._uploadCacheStorageKey());
        return null;
      }
      if (!Array.isArray(parsed.rows) || !parsed.uploadedFilesByFolder || typeof parsed.uploadedFilesByFolder !== 'object') {
        return null;
      }
      return {
        rows: parsed.rows,
        uploadedFilesByFolder: parsed.uploadedFilesByFolder,
      };
    } catch {
      return null;
    }
  }

  private _persistUploadCache(rows: FlatRow[], uploadedFilesByFolder: State['uploadedFilesByFolder']): void {
    if (typeof window === 'undefined') return;
    try {
      const uploadedKeys = Object.keys(uploadedFilesByFolder || {});
      const keysWithFiles = new Set(uploadedKeys.filter(k => Array.isArray(uploadedFilesByFolder[k]) && uploadedFilesByFolder[k].length > 0));
      const cacheRows = rows
        .filter(row => {
          if (!row?.subFolderPath) return false;
          if (row.fileName) return true;
          if (row.groupKey && keysWithFiles.has(row.groupKey)) return true;
          if (row.uploadFolderId && keysWithFiles.has(row.uploadFolderId)) return true;
          return false;
        })
        .slice(-2500);

      if (cacheRows.length === 0 && keysWithFiles.size === 0) {
        window.localStorage.removeItem(this._uploadCacheStorageKey());
        return;
      }

      window.localStorage.setItem(this._uploadCacheStorageKey(), JSON.stringify({
        version: this._uploadCacheVersion,
        savedAt: Date.now(),
        rows: cacheRows,
        uploadedFilesByFolder,
      }));
    } catch {
      // Ignore storage quota/serialization issues.
    }
  }

  private _mergeWithUploadCache(baseRows: FlatRow[], baseUploadedFilesByFolder?: State['uploadedFilesByFolder']): { rows: FlatRow[]; uploadedFilesByFolder: State['uploadedFilesByFolder'] } {
    const cache = this._readUploadCache();
    const mergedUploaded: State['uploadedFilesByFolder'] = { ...(baseUploadedFilesByFolder || {}) };
    const liveRows = this.state.rows.filter(row => row.groupKey.startsWith('live:'));
    const rowsWithLive = [...baseRows, ...liveRows];
    if (!cache) return { rows: rowsWithLive, uploadedFilesByFolder: mergedUploaded };

    const rowSignature = (row: FlatRow): string => [
      (row.vesselName || '').trim().toLowerCase(),
      (row.group || '').trim().toLowerCase(),
      (row.category || '').trim().toLowerCase(),
      (row.subCategory || '').trim().toLowerCase(),
      (row.subFolderPath || '').trim().toLowerCase(),
      (row.fileName || '').trim().toLowerCase(),
      (row.fileId || '').trim().toLowerCase(),
    ].join('||');

    const mergedRows = [...rowsWithLive];
    const seenRows = new Set(mergedRows.map(rowSignature));
    for (const cachedRow of cache.rows) {
      const sig = rowSignature(cachedRow);
      if (!seenRows.has(sig)) {
        mergedRows.push(cachedRow);
        seenRows.add(sig);
      }
    }

    for (const [key, cachedFiles] of Object.entries(cache.uploadedFilesByFolder || {})) {
      if (!Array.isArray(cachedFiles) || cachedFiles.length === 0) continue;
      const existing = mergedUploaded[key] || [];
      const seenNames = new Set(existing.map(file => (file?.name || '').toLowerCase()));
      const mergedFiles = [...existing];
      for (const file of cachedFiles) {
        const name = (file?.name || '').toLowerCase();
        if (!name || seenNames.has(name)) continue;
        seenNames.add(name);
        mergedFiles.push(file);
      }
      mergedUploaded[key] = mergedFiles;
    }

    return { rows: mergedRows, uploadedFilesByFolder: mergedUploaded };
  }

  public _dismissAnomaly = (id: number): void => {
    const base = this._base();
    this._fetchJson(`${base}/api/anomalies/${id}`, {
      method: 'PATCH',
      body: JSON.stringify({ resolved: true }),
    }).catch(() => undefined);
    this.setState(prev => ({
      folderAnomalies: prev.folderAnomalies.filter(a => a.id !== id),
      alertsList: prev.alertsList.filter(a => a.anomaly_id !== id),
    }));
  };

  /**
   * Dismiss anomaly AND add a corresponding entry to the Recycle Bin state
   * so the user can see it in the existing Recycle Bin page.
   */
  public _moveAnomalyToRecycleBin = (anomaly: FolderAnomalyItem): void => {
    // Resolve the anomaly on the backend (mark resolved = true)
    this._dismissAnomaly(anomaly.id);

    // Build a DeletedNode from the anomaly so it appears in Recycle Bin
    const { graphClient: gc, siteId: si, driveId: di } = this.props;
    const useGraph = Boolean(gc && si && di && anomaly.drive_item_id);
    const deletedNode: DeletedNode = {
      id: anomaly.drive_item_id || `anomaly_${anomaly.id}`,
      name: anomaly.name,
      kind: anomaly.item_type === 'file' ? 'file' : 'folder',
      item_type: anomaly.item_type,
      original_path: anomaly.spo_path,
      main_folder: anomaly.department,
      deleted_at: new Date().toISOString(),
      ext: anomaly.item_type === 'file' ? anomaly.name.split('.').pop() : undefined,
      in_spo_recycle_bin: useGraph,
    };

    this.setState(prev => ({
      recycleBin: [...prev.recycleBin, deletedNode],
      spoAnomalyDismissConfirm: null,
    }));

    // Soft-delete via Graph DELETE — moves item to the SPO site Recycle Bin
    if (useGraph) {
      void gc!.api(`/sites/${si}/drives/${di}/items/${anomaly.drive_item_id}`)
        .delete()
        .catch((err: any) => console.warn('[VesselDMS] anomaly soft-delete failed:', err));
    }
  };

  public _fetchAnomalies = (signal?: AbortSignal): void => {
    const base = this._base();
    this._fetchJson(`${base}/api/anomalies`, signal)
      .then(res => {
        if (Array.isArray(res)) {
          this.setState({ folderAnomalies: res }, () => this._syncAnomalyAlerts());
        }
      })
      .catch(() => undefined);
  };

  public _triggerAnomalyScan = (): void => {
    const base = this._base();
    this._fetchJson(`${base}/api/anomalies/scan`, {
      method: 'POST',
    })
      .then(res => {
        if (Array.isArray(res)) {
          this.setState({ folderAnomalies: res }, () => this._syncAnomalyAlerts());
        }
      })
      .catch(() => undefined);
  };

  public _loadNormalFolders = (): void => {
    const base = this._base();
    if (!base) return;
    this._fetchJson(`${base}/api/normal-folders`)
      .then(res => {
        if (Array.isArray(res)) {
          this.setState({ normalFolders: res });
        }
      })
      .catch(() => undefined);
  };

  public _archiveDocumentFile = async (
    fileId: string,
    fileName: string,
    folderPath = '',
    department = '',
    vesselName = '',
  ): Promise<boolean> => {
    if (!fileId || fileId.startsWith('file_') || /^\d+$/.test(fileId)) {
      this.setState({ docUploadMsg: 'Only an approved SharePoint file can be archived.' });
      return false;
    }
    const query = new URLSearchParams({
      type: 'file',
      item_name: fileName,
      department,
      vessel_name: vesselName,
    });
    try {
      const result = await this._fetchJson(`${this._base()}/api/archive/${encodeURIComponent(fileId)}?${query.toString()}`, { method: 'POST' });
      if (result?.status === 'pending') {
        this.setState({ docUploadMsg: `Archive request for "${fileName}" submitted for approval.` });
        return false;
      }
      this.setState(prev => {
        const rows = prev.rows.map(row => row.fileId === fileId ? { ...row, fileName: null, fileId: null, filePending: false } : row);
        const uploadedFilesByFolder = Object.fromEntries(Object.entries(prev.uploadedFilesByFolder).map(([key, files]) => [
          key, files.filter(file => file.id !== fileId && file.name !== fileName),
        ]));
        const nextArchivedFileIds = new Set(prev.archivedFileIds);
        nextArchivedFileIds.add(fileId);
        return {
          rows,
          uploadedFilesByFolder,
          docUploadMsg: `"${fileName}" archived successfully.`,
          listViewSelectedFiles: new Set(Array.from(prev.listViewSelectedFiles).filter(id => id !== fileId)),
          folderViewSelectedFiles: new Set(Array.from(prev.folderViewSelectedFiles).filter(id => id !== fileId)),
          archivedFileIds: nextArchivedFileIds,
        };
      });
      const archiveData = await this._fetchJson(`${this._base()}/api/archive/nodes`);
      this.setState({ archiveList: (archiveData || []).map((item: any) => ({ ...item, name: cleanName(item.name || '') })) });
      void this._loadArchivedFileIds();
      return true;
    } catch (error: any) {
      this.setState({ docUploadMsg: `Archive failed: ${error?.message || 'Could not archive the file.'}` });
      return false;
    }
  };

  public _archiveSelectedDocuments = async (
    files: Array<{ id: string; name: string; folderPath?: string; department?: string; vesselName?: string }>,
  ): Promise<void> => {
    const approvedFiles = files.filter(file => file.id && !file.id.startsWith('file_') && !/^\d+$/.test(file.id));
    if (approvedFiles.length === 0) {
      this.setState({ docUploadMsg: 'Select at least one approved file to archive.' });
      return;
    }
    let archivedCount = 0;
    for (const file of approvedFiles) {
      if (await this._archiveDocumentFile(file.id, file.name, file.folderPath, file.department, file.vesselName)) archivedCount += 1;
    }
    if (archivedCount > 0) {
      await this._goToView('archive');
    }
  };

  /** Opens the Archive popup (see ArchiveSelectionModal) rooted at a given
   *  SharePoint folder — used by the "Archive" toolbar button when browsing
   *  the plain folder tree (Sites / Shared Documents / Documents), which has
   *  no pre-existing checkbox selection to archive the old way. */
  public _openArchivePicker = (
    siteId: string,
    driveId: string,
    folderId: string,
    folderName: string,
    department: string = '',
    vesselName: string = '',
  ): void => {
    if (!siteId || !driveId) {
      this.setState({ docUploadMsg: 'Could not determine which SharePoint folder to archive from. Open a folder first.' });
      return;
    }
    this.setState({
      archivePickerDialog: { siteId, driveId, folderId: folderId || 'root', folderName: folderName || 'this folder', department, vesselName },
    });
  };

  public _renderArchivePickerDialog(): React.ReactElement | null {
    return renderArchivePickerModal(this);
  }

  /** Opens the "New Folder" popup (see CreateFolderModal) rooted at a given
   *  SharePoint folder — used by the Documents toolbar's "New Folder" button
   *  when browsing the plain folder tree (Sites / Shared Documents /
   *  Documents), where folders are arbitrary and not DB-backed the way a
   *  vessel's folders are (see _openAddFolderDialog for that case). */
  public _openCreateFolderDialog = (
    siteId: string,
    driveId: string,
    folderRef: string,
    displayPath: string,
    vesselName: string = '',
  ): void => {
    if (!siteId || !driveId) {
      this.setState({ docUploadMsg: 'Could not determine which SharePoint folder to create the folder in. Open a folder first.' });
      return;
    }
    this.setState({
      createFolderDialog: { siteId, driveId, folderRef: folderRef || 'root', displayPath: displayPath || 'this folder', vesselName },
    });
  };

  /** Creates a single named folder under an arbitrary Graph folder reference
   *  (id, name, or slash/" > "-delimited path) via
   *  POST /api/sites/{site}/drives/{drive}/folders/{folderRef}/create-folder. */
  public _createFolderAtPath = async (
    siteId: string,
    driveId: string,
    folderRef: string,
    name: string,
  ): Promise<{ success: boolean; error?: string }> => {
    const encodedRef = (folderRef || 'root')
      .split('/')
      .map(seg => encodeURIComponent(seg))
      .join('/');
    try {
      const res = await fetch(
        `${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/folders/${encodedRef}/create-folder`,
        {
          method: 'POST',
          headers: this._headers(),
          body: JSON.stringify({ name, user_email: this.props.userEmail || undefined }),
        },
      );
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        return { success: false, error: data?.detail || data?.message || `Error ${res.status}` };
      }
      const parentId = data?.parentReference?.id || 'root';
      void this._refreshSiteFolder(siteId, driveId, parentId).catch(() => undefined);
      return { success: true };
    } catch (error: any) {
      return { success: false, error: error?.message || 'Could not create the folder. Check your connection and try again.' };
    }
  };

  public _renderCreateFolderDialog(): React.ReactElement | null {
    return renderCreateFolderModal(this);
  }

  public _restoreArchivedItem = async (item: DeletedNode): Promise<void> => {
    try {
      const result = await this._fetchJson(`${this._base()}/api/restore/${encodeURIComponent(item.id)}?type=${encodeURIComponent(item.kind === 'file' ? 'file' : 'folder')}&item_name=${encodeURIComponent(item.name)}`, { method: 'POST' });
      if (result?.status === 'pending') {
        this.setState({ docUploadMsg: `Restore request for "${item.name}" submitted for approval.` });
        return;
      }
      this.setState(prev => {
        const nextArchivedFileIds = new Set(prev.archivedFileIds);
        nextArchivedFileIds.delete(item.id);
        return {
          archiveList: prev.archiveList.filter(archived => archived.id !== item.id),
          archivedFileIds: nextArchivedFileIds,
          docUploadMsg: `"${item.name}" restored successfully.`,
        };
      });
      await this._loadData(true);
    } catch (error: any) {
      this.setState({ docUploadMsg: `Restore failed: ${error?.message || 'Could not restore the item.'}` });
    }
  };

  /** Archive page's "Move to Recycle Bin" action — distinct from Restore:
   *  actually deletes the item from SharePoint (into the Recycle Bin) via
   *  POST /api/archive/{id}/recycle, instead of un-archiving it back into
   *  the working Documents view. */
  public _moveArchivedItemToRecycleBin = async (item: DeletedNode): Promise<void> => {
    try {
      const type = item.kind === 'file' ? 'file' : 'folder';
      const result = await this._fetchJson(
        `${this._base()}/api/archive/${encodeURIComponent(item.id)}/recycle?type=${encodeURIComponent(type)}`,
        { method: 'POST' },
      );
      if (result?.status === 'pending') {
        this.setState({ docUploadMsg: `Recycle-bin request for "${item.name}" submitted for approval.` });
        return;
      }
      this.setState(prev => {
        const nextArchivedFileIds = new Set(prev.archivedFileIds);
        nextArchivedFileIds.delete(item.id);
        return {
          archiveList: prev.archiveList.filter(archived => archived.id !== item.id),
          archivedFileIds: nextArchivedFileIds,
          docUploadMsg: `"${item.name}" moved to Recycle Bin.`,
        };
      });
      await this._loadRecycleBin();
    } catch (error: any) {
      this.setState({ docUploadMsg: `Could not move "${item.name}" to the Recycle Bin: ${error?.message || 'unknown error'}` });
    }
  };

  /** Loads (or reloads) the Archive page's list — shared by the Archive
   *  nav item's _goToView('archive') and the page's own mount effect, so
   *  navigating there twice or refreshing it doesn't duplicate this logic. */
  public _loadArchiveList = async (): Promise<void> => {
    this.setState({ panelLoading: true });
    try {
      const data = await this._fetchJson(`${this._base()}/api/archive/nodes`);
      this.setState({ archiveList: (data || []).map((v: any) => ({ ...v, name: cleanName(v.name || '') })), panelLoading: false });
    } catch {
      this.setState({ archiveList: [], panelLoading: false });
    }
  };

  public _saveNormalFolder = async (anomaly: FolderAnomalyItem): Promise<any> => {
    const base = this._base();
    return this._fetchJson(`${base}/api/normal-folders`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        drive_item_id: anomaly.drive_item_id,
        name: anomaly.name,
        item_type: anomaly.item_type,
        spo_path: anomaly.spo_path,
        department: anomaly.department || '',
        vessel_name: anomaly.vessel_name || null,
      }),
    });
  };

  public _placeAnomalyInDmsCategory = async (
    anomaly: FolderAnomalyItem,
    scope: 'vessels' | 'common' | 'kaizen',
    vesselName: string,
    mainFolder: string,
    folderPath: string[],
  ): Promise<any> => {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) throw new Error('SharePoint connection is unavailable.');

    const prefix = scope === 'vessels'
      ? [mainFolder, vesselName]
      : scope === 'common'
        ? [mainFolder, 'Common for all ships']
        : ['Kaizen - Knowledge Bank'];
    const pathParts = [...prefix, ...folderPath].filter(Boolean);
    const targetPathParts = pathParts.filter(Boolean);
    const targetPath = targetPathParts.join('/');
    const normaliseSegment = (value: string): string => value
      .replace(/^folder-\d+\s+/i, '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    const targetSegments = targetPathParts.map(normaliseSegment);
    let target: { id: string; name: string; isFolder: boolean } | undefined;

    // Prefer the delta map because it contains the live IDs and actual SPO names.
    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder || node.deleted) continue;
      const liveSegments = (node.serverRelativePath || '').split('/').filter(Boolean).map(normaliseSegment);
      if (liveSegments.length >= targetSegments.length &&
          liveSegments.slice(-targetSegments.length).join('/') === targetSegments.join('/')) {
        target = { id: node.id, name: node.name, isFolder: true };
        break;
      }
    }

    // Walk the live Graph hierarchy when the delta map has not loaded the branch.
    if (!target) {
      let actualPath = '';
      let currentItems: Array<{ id: string; name: string; isFolder: boolean }> = [];
      for (const segment of targetSegments) {
        currentItems = await this._getGraphChildren(actualPath, new AbortController().signal);
        const match = currentItems.find(item => item.isFolder && normaliseSegment(item.name) === segment);
        if (!match) {
          currentItems = [];
          break;
        }
        actualPath = actualPath ? `${actualPath}/${match.name}` : match.name;
        target = { id: match.id, name: match.name, isFolder: true };
      }
    }
    if (!target?.id) throw new Error(`The destination folder "${targetPath}" was not found in SharePoint.`);

    await graphClient
      .api(`/sites/${siteId}/drives/${driveId}/items/${encodeURIComponent(anomaly.drive_item_id)}`)
      .patch({ parentReference: { id: target.id } });

    return this._saveNormalFolder({
      ...anomaly,
      spo_path: `${targetPath}/${anomaly.name}`,
      department: mainFolder,
      vessel_name: scope === 'vessels' ? vesselName : scope === 'common' ? 'Common for all vessels' : 'Kaizen - Knowledge Bank',
    });
  };




  public componentDidMount(): void {
    this._isUnmounted = false;
    void this._loadDocumentSites().finally(() => {
      if (this.props.sessionId) this.setState({ sessionReady: true });
      this._loadData();
    });
    this._startDeltaSync();
    this._loadBentoConfig();
    void this._loadModuleSettings();
    void this._loadSettingsTabSettings();
    void this._loadFilterSettings();
    void this._loadColorSettings();
    void this._loadDashboardStats();
    void this._loadArchivedFileIds();
    void this._loadAllVesselNamesForSearch();
    this._alertRefreshTimer = setInterval(() => this._fetchAlerts(), 30000);
    this._handleResize();
    window.addEventListener('resize', this._handleResize);
    document.addEventListener('click', this._handleOutsideClick);
    document.addEventListener('keydown', this._handleDocumentNavigationKeyDown);
  }

  public _handleOutsideClick = (e: MouseEvent): void => {
    if (!this.state.alertOpen) return;
    const target = e.target as HTMLElement | null;
    if (!target || !target.closest('[data-alert-bell]')) {
      this._closeAlertBell();
    }
  };

  /** Refresh the set of archived drive item IDs (GET /api/archive/ids) so
   *  Documents/Folder/List views can hide archived files without needing
   *  the full Archive page's data — they only become visible again on the
   *  Archive page, or once restored. Called on mount and after every
   *  archive/restore so the Documents view updates immediately, not just
   *  the Archive page. */
  public _loadArchivedFileIds = async (): Promise<void> => {
    const base = this._base();
    if (!base) return;
    try {
      const ids = await this._fetchJson(`${base}/api/archive/ids`);
      if (this._isUnmounted) return;
      this.setState({ archivedFileIds: new Set(Array.isArray(ids) ? ids.map((id: any) => String(id)) : []) });
    } catch {
      // Non-fatal — Documents/Folder views just won't hide archived files
      // until the next successful refresh.
    }
  };

  public _loadBentoConfig(): void {
    this._fetchJson(`${this._base()}/api/email-notification/config`)
      .then((cfg: any) => {
        const recipient = cfg?.ai_bento_recipient || cfg?.graph_sender_mailbox || '';
        this.setState({ bentoConfigRecipient: recipient });
      })
      .catch(() => undefined);
  }

  /** Settings → Module Management (module_settings_api.py). Loads which
   *  modules are currently hidden app-wide so Sidebar can filter its nav
   *  and _goToView/render can refuse to land on a hidden module. Non-fatal
   *  on failure — the app just shows every module until the next reload. */
  public _loadModuleSettings = async (): Promise<void> => {
    try {
      const data = await this._fetchJson(`${this._base()}/api/module-settings/config`);
      if (this._isUnmounted) return;
      const hidden = Array.isArray(data?.hidden) ? data.hidden.filter((id: unknown) => typeof id === 'string') : [];
      this.setState({ hiddenModules: hidden });
    } catch {
      // Non-fatal — see comment above.
    }
  };

  /** Settings → Settings Management (settings_tab_api.py). Loads which
   *  Settings-page tabs are currently hidden app-wide so SettingsPage can
   *  filter its own left nav and redirect off a tab that just got hidden.
   *  Non-fatal on failure — the app just shows every tab until the next
   *  reload. */
  public _loadSettingsTabSettings = async (): Promise<void> => {
    try {
      const data = await this._fetchJson(`${this._base()}/api/settings-tab-settings/config`);
      if (this._isUnmounted) return;
      const hidden = Array.isArray(data?.hidden) ? data.hidden.filter((id: unknown) => typeof id === 'string') : [];
      this.setState({ hiddenSettingsTabs: hidden });
    } catch {
      // Non-fatal — see comment above.
    }
  };

  /** Settings → Filter Search Management (filter_settings_api.py). Loads
   *  whether the Documents page's filter toolbar should render as the
   *  inline dropdown row or the slide-out Filters panel, app-wide.
   *  Non-fatal on failure — the app just keeps the default ('dropdown')
   *  until the next reload. */
  public _loadFilterSettings = async (): Promise<void> => {
    try {
      const data = await this._fetchJson(`${this._base()}/api/filter-settings/config`);
      if (this._isUnmounted) return;
      const mode = data?.mode === 'panel' ? 'panel' : 'dropdown';
      this.setState({ filterUiMode: mode });
    } catch {
      // Non-fatal — see comment above.
    }
  };

  /** Settings → Color Management (color_settings_api.py). Loads the
   *  app-wide background/text/design/hover colors (light + night mode) and
   *  applies them via applyColorTheme() so every page picks them up
   *  through the clay.* CSS-variable tokens. Non-fatal on failure — the
   *  app just keeps rendering with the Ocean Clay defaults already applied
   *  at construction time. */
  public _loadColorSettings = async (): Promise<void> => {
    try {
      const data = await this._fetchJson(`${this._base()}/api/color-settings/config`);
      if (this._isUnmounted) return;
      const colors: ClayColorTheme | undefined = data && data.colors && data.colors.light && data.colors.night ? data.colors : undefined;
      if (!colors) return;
      applyColorTheme(colors);
      this.setState({ colorTheme: colors });
    } catch {
      // Non-fatal — see comment above.
    }
  };

  public _loadDashboardStats(forceRefresh = false): Promise<void> {
    const base = this._base();
    if (!base) return Promise.resolve();
    const siteFilterValue = this.state.dashboardSiteFilter && this.state.dashboardSiteFilter !== 'all'
      ? this.state.dashboardSiteFilter
      : '';
    const params = new URLSearchParams();
    if (forceRefresh) params.set('force_refresh', 'true');
    if (siteFilterValue) params.set('site_key', siteFilterValue);
    const qs = params.toString();
    const url = `${base}/api/dashboard/stats${qs ? `?${qs}` : ''}`;
    // Two calls to this method can be in flight at once (initial mount, a
    // Settings/Home site change, navigating back to Home, ...), and a slow
    // one — e.g. a full live scan that started before a site was
    // hidden/unhidden/added and only finishes 20-30s later — can resolve
    // AFTER a faster, more recent one. Applying it blindly would overwrite
    // the newer state with stale data. Guard by request order (this seq),
    // not by last_refreshed_epoch: that's a per-site server cache
    // timestamp, and a newly selected site's cache is very often older
    // than the previously selected site's, which was dropping the fresh
    // response for the new site and leaving the old site's stats on
    // screen until a later background refresh happened to produce a
    // higher epoch — the "site only updates after a few seconds" bug.
    const seq = ++this._dashboardStatsSeq;
    return this._fetchJson(url)
      .then((data: any) => {
        if (seq < this._dashboardStatsAppliedSeq) return;
        if (data && typeof data === 'object' && 'total_documents' in data) {
          this._dashboardStatsAppliedSeq = seq;
          this.setState({ dashboardStats: data as import('./pages/DashboardPage').DashboardStats });
          if (Array.isArray(data.sites) && data.sites.some((site: any) => site?.stats_pending)) {
            window.setTimeout(() => {
              if (!this._isUnmounted && seq === this._dashboardStatsSeq) {
                void this._loadDashboardStats(false);
              }
            }, 2000);
          }
        }
      })
      .catch(() => undefined);
  }

  // Home/Dashboard module's own SharePoint-site filter. Kept separate from
  // vesselSiteFilter (Vessels page) and activeDocumentSite (Documents/folder
  // browsing) so switching the site on the Home dashboard doesn't affect
  // those other views. Re-fetches dashboard stats (documents, vessels,
  // recent documents) scoped to the chosen site.
  public _handleDashboardSiteChange = (siteKey: string): void => {
    if (siteKey === this.state.dashboardSiteFilter) return;
    // Read from the backend's per-site cache (CACHE_TTL_DASHBOARD_STATS,
    // 120s) instead of forcing a live Graph re-scan on every dropdown
    // change. force_refresh=true here was making every site switch pay for
    // a full live scan (can take 20-30s per the backend's own comments),
    // even when that site's stats were already cached from a few seconds
    // ago — which is what made switching feel "slow, again and again".
    // Explicit refresh (the Refresh button, _refreshCurrentModule) still
    // passes force_refresh=true on purpose.
    this.setState({ dashboardSiteFilter: siteKey }, () => {
      void this._loadDashboardStats(false);
    });
  };

  public componentDidUpdate(pp: IVesselEmailProps, ps: State): void {
    // Mark session as ready once we receive a valid session_id prop
    if (!this.state.sessionReady && this.props.sessionId) {
      this.setState({ sessionReady: true });
    }
    // Persist the current folder navigation (selected main folder / sub-folder
    // / breadcrumb) so a page refresh can restore exactly where the user was,
    // instead of dropping back to the Documents root. Reference-equality
    // checks are enough here: every place that changes these fields does so
    // via an immutable update (a fresh array/object), never a mutation.
    if (ps.folderPathStack !== this.state.folderPathStack || ps.docMainFolder !== this.state.docMainFolder ||
      ps.docCategoryFilter !== this.state.docCategoryFilter || ps.docSubfolderOtherFilter !== this.state.docSubfolderOtherFilter ||
      ps.docViewMode !== this.state.docViewMode || ps.docScopeType !== this.state.docScopeType ||
      ps.activeDocumentSite !== this.state.activeDocumentSite) {
      this._persistFolderNav();
    }
    // Trigger initial load when sessionId first arrives (was missing on mount)
    if (!pp.sessionId && this.props.sessionId) {
      this._filesLoadedForVessels.clear();
      this._filesLoadedForFolders.clear();
      this._loadData();
      return;
    }
    // Trigger initial load when session init settles WITHOUT a sessionId
    // (stub/no-DB backend) — otherwise componentDidMount's 2s fallback is the
    // only chance to load, and only if timing lines up.
    if (!pp.sessionInitialized && this.props.sessionInitialized && !this.props.sessionId && this.state.rows.length === 0) {
      this._loadData();
      return;
    }
    // ONLY reload data when the API URL changes or an explicit reload is requested.
    // Do NOT reload on any other state changes – that would create infinite loops.
    if (pp.apiBaseUrl !== this.props.apiBaseUrl) {
      this._filesLoadedForVessels.clear();
      this._filesLoadedForFolders.clear();
      this._loadData();
    } else if (ps.reloadKey !== this.state.reloadKey) {
      this._filesLoadedForVessels.clear();
      this._filesLoadedForFolders.clear();
      this._loadData();
    }
    // Reset pagination when filters change
    if (ps.textFilter !== this.state.textFilter || ps.vesselFilter !== this.state.vesselFilter ||
      ps.groupFilter !== this.state.groupFilter || ps.catFilter !== this.state.catFilter) {
      // Only call setState if page is not already 0 to avoid unnecessary re-render
      if (this.state.page !== 0) this.setState({ page: 0 });
    }
    // On-demand vessel row loading when user selects a specific vessel in the Documents list view.
    // Call the vessel-specific flat-tree endpoint and refresh live SharePoint files.
    // Skipped in the live-library scopes ('sites' / 'shared_docs' / 'documents'):
    // these two calls resolve a vessel's DB rows via the legacy default
    // department template (Technical & Crewing/Vessels/<vessel>, etc.), which
    // doesn't match a real site's actual folder names and just spams Graph
    // with 404s (that's what selecting a vessel while browsing an actual
    // SharePoint site's live folder tree was doing — see the "Technical"
    // chip 404 storm). Live-scope file listings already come from
    // _getOrLoadSiteFolderChildren / the folder tree, not from DB rows.
    const { vesselFilter } = this.state;
    const isLiveLibraryScope = this.state.docScopeType === 'sites' ||
      this.state.docScopeType === 'shared_docs' || this.state.docScopeType === 'documents';
    if (!isLiveLibraryScope && ps.vesselFilter !== vesselFilter && vesselFilter !== 'all') {
      this._lastRefreshedRowsKey = '';
      this._loadVesselRowsFromApi(vesselFilter).catch(() => undefined);
      setTimeout(() => this._refreshFilesFromBackendRows(), 100);
    }
    // A different vessel (or vessel cleared) makes Compare mode's per-main-
    // folder browsing state stale — it was keyed off the previous vessel's
    // folders. Drop it and fall back out of Compare mode rather than showing
    // boxes that no longer correspond to the newly selected vessel.
    if (ps.vesselFilter !== vesselFilter) {
      if (this.state.docCompareMode || Object.keys(this.state.compareBoxSubPaths).length > 0) {
        this.setState({ docCompareMode: false, compareBoxSubPaths: {} });
      }
    }
    if (ps.docListPage !== this.state.docListPage && this.state.docViewMode === 'list') {
      this._lastRefreshedRowsKey = '';
      this._refreshFilesFromBackendRows(this.state.docListPage);
    }

    if (ps.rows !== this.state.rows || ps.uploadedFilesByFolder !== this.state.uploadedFilesByFolder) {
      this._persistUploadCache(this.state.rows, this.state.uploadedFilesByFolder);
    }

    if (ps.pendingVesselSuggestions !== this.state.pendingVesselSuggestions) {
      this._persistPendingVesselSuggestions(this.state.pendingVesselSuggestions || []);
    }
  }

  public componentWillUnmount(): void {
    this._isUnmounted = true;
    this._cancelDocumentLiveTree();
    this._abort?.abort();
    this._syncScheduler?.stop();
    if (this._deltaReloadTimer) clearTimeout(this._deltaReloadTimer);
    if (this._alertRefreshTimer) clearInterval(this._alertRefreshTimer);
    if (this._deleteAutoCloseTimer) clearInterval(this._deleteAutoCloseTimer);
    if (this._vesselToastAutoCloseTimer) clearTimeout(this._vesselToastAutoCloseTimer);
    if (this._documentToastAutoCloseTimer) clearTimeout(this._documentToastAutoCloseTimer);
    window.removeEventListener('resize', this._handleResize);
    document.removeEventListener('click', this._handleOutsideClick);
    document.removeEventListener('keydown', this._handleDocumentNavigationKeyDown);
  }

  public _clearDeleteAutoCloseTimer = (): void => {
    if (this._deleteAutoCloseTimer) {
      clearInterval(this._deleteAutoCloseTimer);
      this._deleteAutoCloseTimer = null;
    }
  };

  public _startDeleteAutoClose = (seconds: number = 10): void => {
    this._clearDeleteAutoCloseTimer();
    this.setState({ deleteAutoCloseSeconds: seconds });
    this._deleteAutoCloseTimer = setInterval(() => {
      this.setState(prev => {
        const current = prev.deleteAutoCloseSeconds ?? seconds;
        if (current <= 1) {
          this._clearDeleteAutoCloseTimer();
          this._closeModal();
          return { deleteAutoCloseSeconds: null };
        }
        return { deleteAutoCloseSeconds: current - 1 };
      });
    }, 1000);
  };

  /** Schedule (or reschedule) the vessel-deletion popup's auto-dismiss.
   *  Called whenever the popup's content changes (new item, or attribution
   *  resolving in) so a slow-arriving update doesn't cut the timer short. */
  public _scheduleVesselToastAutoClose = (): void => {
    if (this._vesselToastAutoCloseTimer) clearTimeout(this._vesselToastAutoCloseTimer);
    this._vesselToastAutoCloseTimer = window.setTimeout(() => {
      this._vesselToastAutoCloseTimer = null;
      this.setState({ spoVesselDeletedToast: null });
    }, VesselEmail._TOAST_AUTO_DISMISS_MS);
  };

  /** Pause the vessel-deletion popup's auto-dismiss while the pointer is
   *  over it (resumed by _scheduleVesselToastAutoClose on mouse-leave). */
  public _pauseVesselToastAutoClose = (): void => {
    if (this._vesselToastAutoCloseTimer) {
      clearTimeout(this._vesselToastAutoCloseTimer);
      this._vesselToastAutoCloseTimer = null;
    }
  };

  /** Same as _scheduleVesselToastAutoClose, for the document/folder
   *  deletion popup (which can accumulate several items into one toast). */
  public _scheduleDocumentToastAutoClose = (): void => {
    if (this._documentToastAutoCloseTimer) clearTimeout(this._documentToastAutoCloseTimer);
    this._documentToastAutoCloseTimer = window.setTimeout(() => {
      this._documentToastAutoCloseTimer = null;
      this.setState({ spoDocumentDeletedToast: null });
    }, VesselEmail._TOAST_AUTO_DISMISS_MS);
  };

  public _pauseDocumentToastAutoClose = (): void => {
    if (this._documentToastAutoCloseTimer) {
      clearTimeout(this._documentToastAutoCloseTimer);
      this._documentToastAutoCloseTimer = null;
    }
  };

  public _handleDocumentNavigationKeyDown = (event: KeyboardEvent): void => {
    if (this.state.view !== 'list' || event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
    const target = event.target as HTMLElement | null;
    if (target?.closest('input, select, textarea, button, [contenteditable="true"]')) return;

    const offset = event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowRight' ? 1 : 0;
    if (!offset) return;
    const nextIndex = this.state.folderNavIndex + offset;
    const next = this.state.folderNavHistory[nextIndex];
    if (!next) return;
    event.preventDefault();
    this.setState({
      folderNavIndex: nextIndex,
      folderPathStack: next.folderPathStack,
      docMainFolder: next.docMainFolder,
    });
  };

  // ── Delta Sync ────────────────────────────────────────────────────────────

  public _startDeltaSync(): void {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    this._syncScheduler = createSyncScheduler(
      graphClient,
      siteId,
      driveId,
      (result: DeltaSyncResult) => this._applyDeltaResult(result),
      60_000,
      (err) => console.warn('[VesselDMS] Delta sync error:', err),
    );
    this._syncScheduler.start();
    // Run immediately on mount
    void this._syncScheduler.triggerNow().catch(() => undefined);
  }

  /**
   * Called when the delta sync reports a deleted SPO item ID that matches a
   * known vessel's top-level folder. Moves the vessel to the recycle bin and
   * emits a vessel_deleted alert in the header bell.
   */
  public _handleSpoVesselDeletion(deletedId: string, vesselFolderNode: import('./deltaSync').SpoFolderNode): boolean {
    const { vessels } = this.state;
    const normName = (s: string): string => cleanName(s).trim().toLowerCase();

    // Guard: only treat as a vessel deletion if the folder was at the vessel level
    // (its path must contain a 'vessels' segment as a parent, not be a sub-folder).
    const pathSegments = (vesselFolderNode.serverRelativePath || '')
      .split('/')
      .map(s => s.trim().toLowerCase())
      .filter(Boolean);
    const vesselsIdx = pathSegments.lastIndexOf('vessels');
    // The folder must be the direct child of a 'vessels' (or 'specific vessels') container.
    const parentSeg = pathSegments[pathSegments.length - 2] || '';
    const isVesselLevel = vesselsIdx !== -1 &&
      (pathSegments.length - 1 === vesselsIdx + 1 ||
        parentSeg === 'specific vessels' ||
        parentSeg === 'vessels');
    if (!isVesselLevel) return false;

    // Already raised a popup for this exact SPO item (this handler can be
    // reached more than once for the same delete via retries/rescans) —
    // report it as "handled" so the caller doesn't fall back to the
    // generic document-deletion popup for the same folder.
    if (this._shownDeletionToastIds.has(deletedId)) return true;

    // Match by folder name against known vessels
    const vessel = vessels.find(v => normName(v.name) === normName(vesselFolderNode.name));
    if (!vessel) return false;

    const now = new Date().toISOString();
    const activeSiteObj = (this.state.documentSites || []).find(s => s.site_key === this.state.activeDocumentSite);
    const siteNames = (vessel.provisioned_site_ids || []).map(sk => {
      const matched = (this.state.documentSites || []).find(s => s.site_key === sk);
      return matched?.sp_site_name || sk;
    }).filter(Boolean);
    const resolvedSiteName = siteNames.length > 0
      ? siteNames.join(', ')
      : (activeSiteObj?.sp_site_name || this.props.siteUrl || 'SharePoint');

    const recycleBinEntry: DeletedNode = {
      id: vessel.id,
      name: vessel.name,
      kind: 'vessel',
      item_type: 'vessel',
      main_folder: this.MAIN_FOLDER_NAMES[0],
      original_path: `${this.MAIN_FOLDER_NAMES[0]}/${vessel.name}`,
      vessel_name: '',
      deleted_at: now,
      imo: vessel.imo,
      vessel_type: vessel.vessel_type,
      site_name: resolvedSiteName,
      site_key: (vessel.provisioned_site_ids || [])[0] || this.state.activeDocumentSite || undefined,
    };

    const alert: import('./types/ui').AlertItem = {
      id: `vessel_deleted_spo_${vessel.id}_${Date.now()}`,
      drive_item_id: deletedId,
      folder_name: vessel.name,
      folder_path: vesselFolderNode.serverRelativePath || recycleBinEntry.original_path || '',
      parent_folder_id: null,
      vessel_name: vessel.name,
      department: 'All Departments',
      created_by_email: '',
      created_by_name: 'Unknown user',
      alert_type: 'vessel_deleted',
      alert_category: 'dms',
      read: false,
      created_at: now,
      item_type: 'folder',
      spo_path: vesselFolderNode.serverRelativePath,
      site_name: resolvedSiteName,
    };

    this._shownDeletionToastIds.add(deletedId);
    this.setState(prev => ({
      vessels: prev.vessels.filter(v => v.id !== vessel.id),
      recycleBin: [
        recycleBinEntry,
        ...prev.recycleBin.filter(r => r.name.toLowerCase() !== vessel.name.toLowerCase()),
      ],
      alertsList: [alert, ...prev.alertsList],
      spoVesselDeletedToast: {
        vesselName: vessel.name,
        id: alert.id,
        siteName: resolvedSiteName,
        originalPath: recycleBinEntry.original_path,
        deletedByName: null,
        deletedByEmail: null,
      },
      spoDeletedVesselIds: (() => { const s = new Set(this.state.spoDeletedVesselIds); s.add(vessel.id); return s; })(),
    }));
    this._scheduleVesselToastAutoClose();

    // Who actually deleted it in SharePoint isn't in the delta-sync event —
    // ask the backend's recycle-bin feed (enriched from SharePoint's own
    // recycle bin once that propagates) and patch the alert + toast in place.
    void this._lookupDeletionAttribution(vessel.id, vessel.name).then(resolved => {
      if (!resolved) return;
      this.setState(prev => ({
        alertsList: prev.alertsList.map(a =>
          a.id === alert.id
            ? {
                ...a,
                created_by_email: resolved.deletedByEmail || a.created_by_email,
                created_by_name: resolved.deletedByName || a.created_by_name,
                site_name: resolved.siteName || a.site_name,
              }
            : a,
        ),
        spoVesselDeletedToast: prev.spoVesselDeletedToast && prev.spoVesselDeletedToast.id === alert.id
          ? {
              ...prev.spoVesselDeletedToast,
              siteName: resolved.siteName || prev.spoVesselDeletedToast.siteName,
              deletedByName: resolved.deletedByName || prev.spoVesselDeletedToast.deletedByName,
              deletedByEmail: resolved.deletedByEmail || prev.spoVesselDeletedToast.deletedByEmail,
            }
          : prev.spoVesselDeletedToast,
      }));
      if (this.state.spoVesselDeletedToast) this._scheduleVesselToastAutoClose();
    });

    // Notify backend so the DB record is soft-deleted (retry once on failure)
    const doDelete = (): Promise<any> =>
      this._fetchJson(`${this._base()}/api/vessels/${vessel.id}?vessel_name=${encodeURIComponent(vessel.name)}`, {
        method: 'DELETE',
      });
    doDelete().catch(() => setTimeout(() => doDelete().catch(() => undefined), 3000));
    return true;
  }

  /**
   * Called when delta sync detects a new folder in SharePoint that is not
   * already known to the DMS. Emits a folder classification alert.
   */
  public _handleSpoNewVesselFolder(node: import('./deltaSync').SpoFolderNode): void {
    const { vessels, folderAnomalies, normalFolders, alertsList } = this.state;
    const normName = (s: string): string => cleanName(s).trim().toLowerCase();

    // Skip structural container folders and known DMS scopes — these are never unrecognised vessels
    const nodeNameNorm = normName(node.name);
    const pathNorm = (node.serverRelativePath || '').toLowerCase();
    if (
      nodeNameNorm === 'specific vessels' ||
      nodeNameNorm === 'vessels' ||
      nodeNameNorm === 'vessel management' ||
      nodeNameNorm === 'common for all ships' ||
      nodeNameNorm === 'common for all vessels' ||
      nodeNameNorm === 'common (not ship specific)' ||
      nodeNameNorm === 'kaizen - knowledge bank' ||
      nodeNameNorm === 'kaizen' ||
      nodeNameNorm === 'knowledge bank' ||
      pathNorm.includes('common for all ships') ||
      pathNorm.includes('common for all vessels') ||
      pathNorm.includes('common (not ship specific)') ||
      pathNorm.includes('kaizen - knowledge bank') ||
      pathNorm.includes('/kaizen')
    ) {
      return;
    }
    // Skip if already a known vessel
    if (vessels.some(v => normName(v.name) === nodeNameNorm)) {
      return;
    }
    // Skip if already classified as a normal folder
    if ((normalFolders || []).some(f => normName(f.name) === nodeNameNorm)) {
      return;
    }
    // Skip if already in anomalies
    if (folderAnomalies.some(a => normName(a.name) === nodeNameNorm)) {
      return;
    }
    // Skip if already in alerts
    if (alertsList.some(a => a.alert_type === 'vessel_unrecognised' && normName(a.folder_name) === nodeNameNorm)) {
      return;
    }

    const now = new Date().toISOString();
    const anomalyId = Date.now() + Math.floor(Math.random() * 100000);
    const anomaly: import('./types/ui').FolderAnomalyItem = {
      id: anomalyId,
      drive_item_id: node.id,
      name: node.name,
      item_type: 'folder',
      anomaly_type: 'vessel_level_unmatched',
      department: 'Vessels',
      vessel_name: null,
      spo_path: node.serverRelativePath,
      resolved: false,
      detected_at: now,
    };

    const alert: import('./types/ui').AlertItem = {
      id: `vessel_unrecognised_${node.id}_${Date.now()}`,
      drive_item_id: node.id,
      folder_name: node.name,
      folder_path: node.serverRelativePath,
      parent_folder_id: node.parentId,
      vessel_name: null,
      department: 'Vessels',
      created_by_email: '',
      created_by_name: 'SharePoint Online',
      alert_type: 'vessel_unrecognised',
      read: false,
      created_at: now,
      anomaly_id: anomalyId,
      item_type: 'folder',
      spo_path: node.serverRelativePath,
    };

    // Update anomalies and alerts without interrupting user workflow or forcing open dialog
    this.setState(prev => ({
      folderAnomalies: [...prev.folderAnomalies, anomaly],
      alertsList: [alert, ...prev.alertsList],
    }));
  }

  /**
   * After the baseline delta sync, scan the populated spoFolderMap for
   * vessel-level folders not registered in the DB. Called once, 5s after
   * the first delta result, to give _loadData time to populate vessels state.
   */
  public _scanSpoFolderMapForNewVessels(): void {
    const { spoFolderMap, vessels } = this.state;
    const normName = (s: string): string => cleanName(s).trim().toLowerCase();
    const knownVesselNames = new Set(vessels.map(v => normName(v.name)));

    for (const [, node] of Array.from(spoFolderMap.entries())) {
      if (!node.isFolder || node.deleted) continue;
      const segs = (node.serverRelativePath || '')
        .split('/').map(s => s.trim().toLowerCase()).filter(Boolean);
      const parentSeg = segs[segs.length - 2] || '';
      const KNOWN_MAIN_DEPTS = ['technical & crewing', 'commercial & chartering', 'insurance'];
      // Accept vessel folders directly under a main department OR under legacy Specific Vessels
      if (parentSeg !== 'specific vessels' && parentSeg !== 'vessels' && !KNOWN_MAIN_DEPTS.includes(parentSeg)) continue;
      // Skip the container folders and common/kaizen roots themselves
      const nameNorm = normName(node.name);
      if (
        nameNorm === 'specific vessels' ||
        nameNorm === 'vessels' ||
        nameNorm === 'vessel management' ||
        nameNorm === 'common for all ships' ||
        nameNorm === 'common for all vessels' ||
        nameNorm === 'common (not ship specific)' ||
        nameNorm === 'kaizen - knowledge bank' ||
        nameNorm === 'kaizen' ||
        nameNorm === 'knowledge bank'
      ) continue;

      this._handleSpoNewVesselFolder(node);
    }
  }

  /**
   * Called when delta sync detects a new file added in SharePoint.
   * Emits an alert if the file is outside the DMS structure.
   */
  public _handleSpoNewFile(node: import('./deltaSync').SpoFolderNode): void {
    const { vessels, alertsList } = this.state;
    const normName = (s: string): string => cleanName(s).trim().toLowerCase();

    if (this._appUploadedFileIds.has(node.id)) {
      this._appUploadedFileIds.delete(node.id);
      return;
    }

    // Skip if file is located inside known DMS scopes like Kaizen or Common for all ships/vessels
    const pathNorm = (node.serverRelativePath || '').toLowerCase();
    if (
      pathNorm.includes('common for all ships') ||
      pathNorm.includes('common for all vessels') ||
      pathNorm.includes('kaizen - knowledge bank') ||
      pathNorm.includes('/kaizen')
    ) {
      return;
    }

    // Skip if already alerted for this file
    if (alertsList.some(a => a.drive_item_id === node.id && a.alert_type === 'file_outside_structure')) return;
    // Determine which vessel this file belongs to (if any) by walking path segs
    const segs = (node.serverRelativePath || '')
      .split('/').map(s => s.trim()).filter(Boolean);
    // Find vessel name: the segment right after 'specific vessels' or 'vessels'
    let vesselName: string | null = null;
    for (let i = 0; i < segs.length - 1; i++) {
      const s = segs[i].toLowerCase();
      const KNOWN_MAIN_DEPTS_LOWER = ['technical & crewing', 'commercial & chartering', 'insurance'];
      if (s === 'specific vessels' || s === 'vessels' || KNOWN_MAIN_DEPTS_LOWER.includes(s)) {
        const candidate = segs[i + 1];
        if (candidate && vessels.some(v => normName(v.name) === normName(candidate))) {
          vesselName = candidate;
        }
        break;
      }
    }

    // If file is inside a known subcategory (depth >= 4 segments under specific vessels), it's not outside structure
    if (vesselName && segs.length >= 5) {
      return;
    }

    const now = new Date().toISOString();
    const alert: import('./types/ui').AlertItem = {
      id: `file_outside_${node.id}_${Date.now()}`,
      drive_item_id: node.id,
      folder_name: node.name,
      folder_path: node.serverRelativePath,
      parent_folder_id: node.parentId,
      vessel_name: vesselName,
      department: 'Vessels',
      created_by_email: '',
      created_by_name: 'SharePoint Online',
      alert_type: 'file_outside_structure',
      read: false,
      created_at: now,
      item_type: 'file',
      spo_path: node.serverRelativePath,
    };

    // Update alerts list in background without interrupting user with popup
    this.setState(prev => ({
      alertsList: [alert, ...prev.alertsList],
    }));

    // NOTE: this used to eagerly call _loadVesselRowsFromApi(vesselName) here
    // ("so the data is ready when user views alerts") — but _handleSpoNewFile
    // fires once per file the site-wide delta sync reports, for whichever
    // vessel that file happens to belong to, completely independent of what
    // the user is currently looking at (this component is a single mounted
    // instance shared across every module/page, so this ran the same way
    // whether the user was on Shared Documents, the Dashboard, or Vessels).
    // On a normal incremental poll that can be a handful of vessels; right
    // after a resync (or the first poll of a long-lived browser tab, since
    // the deltaLink persists in localStorage across page loads) it can be
    // dozens fired within the same tick, each pulling flat-tree rows and
    // then recursively walking that vessel's SharePoint folders for files
    // (_refreshFilesFromBackendRows / _mergeLiveSharePointFiles) — the
    // "why is it loading Belle Lune while I'm on Shared Documents" bug, and
    // the request-volume/throttling behind the failing children/delta calls
    // and the app's public-origin → http://127.0.0.1:8000 fallback firing
    // per failed request (see _fetchJson), which is what the browser flags
    // as a blocked/CORB-tripping local-network request.
    //
    // The alert above is what actually needs to exist immediately (it's a
    // cheap state update, no network call); the vessel's rows only need to
    // exist once the user actually opens that vessel or clicks the alert,
    // which already loads on demand via componentDidUpdate's vesselFilter
    // handler and FolderView's own click-to-open loading. So no fetch here.
  }

  public _applyDeltaResult(result: DeltaSyncResult): void {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    // What SharePoint changed directly, keyed by stable Graph item id (never
    // by name), so the Documents lists can follow it:
    //  goneIds      deleted file, or deleted folder and everything under it
    //  movedFileIds files whose parent folder changed
    //  renamedFiles file id -> new name
    //  changedNodes renamed/moved files and folders (for the folder listings)
    const goneIds = new Set<string>();
    const movedFileIds = new Set<string>();
    const renamedFiles = new Map<string, string>();
    const changedNodes: Array<{ node: import('./deltaSync').SpoFolderNode; renamed: boolean; moved: boolean }> = [];

    this.setState(prev => {
      const map = new Map(prev.spoFolderMap);
      goneIds.clear(); movedFileIds.clear(); renamedFiles.clear(); changedNodes.length = 0;
      const collectGone = (id: string): void => {
        if (goneIds.has(id)) return;
        goneIds.add(id);
        const n = map.get(id);
        if (n) n.children.forEach(c => collectGone(c.id));
      };

      // Process deletions first — check for vessel folder deletions before removing from map
      for (const id of result.deleted) {
        const node = map.get(id);
        const appOriginated = this._appDeletedItemIds.delete(id);
        if (node?.isFolder && !appOriginated) {
          // Defer vessel deletion handling to after setState (needs current vessels state)
          const isVesselFolder = (node.serverRelativePath || '').toLowerCase().includes('/specific vessels/') &&
            (node.serverRelativePath || '').split('/').filter(Boolean).length <= 4;
          if (isVesselFolder) {
            setTimeout(() => {
              if (!this._handleSpoVesselDeletion(id, node)) this._handleSpoDocumentDeletion(node);
            }, 0);
          } else {
            setTimeout(() => this._handleSpoDocumentDeletion(node), 0);
          }
        } else if (node && !appOriginated) {
          setTimeout(() => this._handleSpoDocumentDeletion(node), 0);
        } else if (!node && !appOriginated) {
          // When the deleted item wasn't present in the local map (e.g., a
          // fast external delete before branch materialized), still emit an
          // immediate UI alert/popup by resolving from recycle-bin metadata.
          setTimeout(() => this._handleUnknownSpoDeletion(id), 0);
        }
        collectGone(id);
        removeNodeFromMap(map, id);
      }

      // Process additions/updates — detect new vessel-level folders
      const missingParentIds = new Set<string>();
      for (const node of result.added) {
        // Skip nodes we locally deleted — Graph delta may lag before reporting deletion
        if (this._appDeletedItemIds.has(node.id)) continue;

        // Snapshot before merging: the merge updates the stored node in place.
        const before = map.get(node.id);
        const prevName = before ? before.name : null;
        const prevParentId = before ? before.parentId : null;
        const prevPath = before ? before.serverRelativePath : '';

        const { missingParentId } = mergeNodeIntoMap(map, node);
        if (missingParentId) missingParentIds.add(missingParentId);

        // Same id, different name or parent => renamed / moved in SharePoint,
        // not a new item.
        let renamed = false;
        let moved = false;
        if (before && !result.isBaseline) {
          renamed = prevName !== null && prevName !== node.name;
          moved = prevParentId !== node.parentId;
          if (renamed || moved) {
            changedNodes.push({ node, renamed, moved });
            if (node.isFolder) {
              repathDescendants(map.get(node.id)!, prevPath, node.serverRelativePath);
            } else if (moved) {
              movedFileIds.add(node.id);
            } else {
              renamedFiles.set(node.id, node.name);
            }
            console.info('[VesselDMS] SharePoint change applied', { id: node.id, kind: node.isFolder ? 'folder' : 'file', from: prevName, to: node.name, moved });
          }
        }
        if (renamed || moved) {
          // A renamed/moved item is not new: skip new-folder / new-file handling.
          if (!node.isFolder && moved && renamed) renamedFiles.set(node.id, node.name);
          continue;
        }

        // Detect new folders added anywhere in SPO.
        // Skip during baseline scan so existing folders do not create alerts.
        if (node.isFolder && !result.isBaseline) {
          // Defer to after setState so we have the latest folder and vessel state.
          setTimeout(() => this._handleSpoNewVesselFolder(node), 0);
        }
        // Detect files uploaded anywhere in the site so they can be classified.
        if (!node.isFolder && !result.isBaseline) {
          setTimeout(() => this._handleSpoNewFile(node), 0);
        }
      }

      // For any node whose parent branch isn't loaded, trigger a targeted re-fetch
      // (done outside setState to avoid async inside setState)
      if (missingParentIds.size > 0) {
        setTimeout(() => {
          missingParentIds.forEach(parentId => {
            const parentNode = map.get(parentId);
            if (parentNode) {
              fetchFolderChildren(graphClient!, siteId!, driveId!, parentNode.serverRelativePath)
                .then(children => {
                  this.setState(prev2 => {
                    const m2 = new Map(prev2.spoFolderMap);
                    children.forEach(c => mergeNodeIntoMap(m2, c));
                    return { spoFolderMap: m2, lastDeltaSync: new Date() };
                  });
                })
                .catch(() => undefined);
            }
          });
        }, 0);
      }

      // Direct file mapping from delta results into rows and uploadedFilesByFolder:
      const updatedByFolder = { ...prev.uploadedFilesByFolder };
      let updatedRows = [...prev.rows];
      const addedFiles = result.added.filter(n => !n.isFolder && n.parentId);

      // Apply SharePoint-side delete / rename / move to the Documents lists
      // before re-adding files, matched by item id. A moved file is removed
      // from its old folder here and re-added under its new parent below.
      if (goneIds.size > 0 || movedFileIds.size > 0 || renamedFiles.size > 0) {
        const dropIds = new Set<string>(Array.from(goneIds).concat(Array.from(movedFileIds)));
        for (const key of Object.keys(updatedByFolder)) {
          if (goneIds.has(key)) { delete updatedByFolder[key]; continue; }
          const current = updatedByFolder[key] as any[];
          let touched = false;
          const next: any[] = [];
          for (const f of current) {
            const fid = String(f?.id ?? '');
            if (dropIds.has(fid)) { touched = true; continue; }
            const newName = renamedFiles.get(fid);
            if (newName !== undefined && f.name !== newName) { touched = true; next.push({ ...f, name: newName }); continue; }
            next.push(f);
          }
          if (touched) updatedByFolder[key] = next;
        }
        updatedRows = updatedRows.map(r => {
          const fid = r.fileId ? String(r.fileId) : '';
          if (!fid) return r;
          if (dropIds.has(fid)) return { ...r, fileName: null, fileId: null };
          const newName = renamedFiles.get(fid);
          return newName !== undefined && r.fileName !== newName ? { ...r, fileName: newName } : r;
        });
      }

      // Same for the cached SharePoint folder listings (SharePoint Sites view).
      if (goneIds.size > 0 || changedNodes.length > 0) {
        this._siteFolderItemsCache.forEach((entry, key) => {
          let items: any[] = entry.items as any[];
          let touched = false;
          if (goneIds.size > 0) {
            const kept = items.filter(it => !goneIds.has(String(it?.id ?? '')));
            if (kept.length !== items.length) { items = kept; touched = true; }
          }
          for (const c of changedNodes) {
            const idx = items.findIndex(it => it?.id === c.node.id);
            if (idx < 0) continue;
            if (c.moved) { items = items.filter((_, i) => i !== idx); touched = true; }
            else if (c.renamed) { items = items.map((it, i) => (i === idx ? { ...it, name: c.node.name } : it)); touched = true; }
          }
          for (const c of changedNodes) {
            if (c.moved && c.node.parentId && key.endsWith(`::${c.node.parentId}`) && !items.some(it => it?.id === c.node.id)) {
              items = items.concat([{
                id: c.node.id, name: c.node.name,
                ...(c.node.isFolder ? { folder: { childCount: 0 } } : { file: {} }),
                size: c.node.size, lastModifiedDateTime: c.node.lastModifiedDateTime, webUrl: '',
              }]);
              touched = true;
            }
          }
          if (touched) this._siteFolderItemsCache.set(key, { ...entry, items } as any);
        });
      }

      if (addedFiles.length > 0) {
        for (const fileNode of addedFiles) {
          const parentNode = map.get(fileNode.parentId!);
          const parentPath = parentNode?.serverRelativePath || '';
          const uploadedAt = fileNode.createdDateTime
            ? Date.parse(fileNode.createdDateTime)
            : (fileNode.lastModifiedDateTime ? Date.parse(fileNode.lastModifiedDateTime) : undefined);
          const newFile = {
            name: fileNode.name,
            size: '—',
            date: uploadedAt ? new Date(uploadedAt).toLocaleString() : 'Today',
            pending: false,
            id: fileNode.id,
            uploadedAt,
          };

          // 1. Store under parent folder ID and parent path
          const curParentFiles = updatedByFolder[fileNode.parentId!] || [];
          if (!curParentFiles.some((f: any) => f.name === fileNode.name)) {
            updatedByFolder[fileNode.parentId!] = [...curParentFiles, newFile];
          }
          const normalise = (p: string) => (p || '').replace(/^\/+/, '').replace(/\/+$/, '').toLowerCase();
          const normParent = normalise(parentPath);
          if (normParent) {
            const curNormFiles = updatedByFolder[normParent] || [];
            if (!curNormFiles.some((f: any) => f.name === fileNode.name)) {
              updatedByFolder[normParent] = [...curNormFiles, newFile];
            }
          }

          // 2. Find and update matching rows in updatedRows
          updatedRows.forEach((r, idx) => {
            const matchesId = r.uploadFolderId === fileNode.parentId!;
            let matchesPath = false;
            if (r.subFolderPath) {
              const liveId = this._getLiveSharePointFolderId(r.subFolderPath);
              if (liveId && liveId === fileNode.parentId) {
                matchesPath = true;
              } else if (normParent) {
                const normSub = normalise(this._sharePointFolderPath(r.subFolderPath, ''));
                if (normSub && (normParent === normSub || normParent.endsWith(`/${normSub}`))) {
                  matchesPath = true;
                }
              }
            }

            if (matchesId || matchesPath) {
              const normSub = (r.subFolderPath || '').trim().toLowerCase();
              if (normSub) {
                const curSubFiles = updatedByFolder[normSub] || [];
                if (!curSubFiles.some((f: any) => f.name === fileNode.name)) {
                  updatedByFolder[normSub] = [...curSubFiles, newFile];
                }
              }
              if (r.groupKey) {
                const curGFiles = updatedByFolder[r.groupKey] || [];
                if (!curGFiles.some((f: any) => f.name === fileNode.name)) {
                  updatedByFolder[r.groupKey] = [...curGFiles, newFile];
                }
              }
              if (!r.fileName) {
                updatedRows[idx] = {
                  ...r,
                  fileName: fileNode.name,
                  fileId: fileNode.id,
                  filePending: false,
                };
              }
            }
          });
        }
      }

      return { spoFolderMap: map, lastDeltaSync: new Date(), rows: updatedRows, uploadedFilesByFolder: updatedByFolder };
    }, () => {
      // After the baseline delta sync populates spoFolderMap, scan it for
      // vessel-level folders that are not in the DB. Deferred 5s to ensure
      // _loadData has finished populating the vessels state first.
      if (result.isBaseline) {
        setTimeout(() => this._scanSpoFolderMapForNewVessels(), 5000);
      }
      // The delta map is the source of truth for folders in this Documents
      // drive. On first delta result, kick off a spoFolderMap-based file
      // refresh (fast: only queries confirmed live folder IDs).
      if (!this._initialDocumentFolderRefreshDone && this.state.spoFolderMap.size > 0 && this.state.rows.length > 0) {
        this._initialDocumentFolderRefreshDone = true;
        // Reset guard so the refresher picks up the now-populated spoFolderMap
        this._lastRefreshedRowsKey = '';
        this._refreshFilesFromBackendRows();
      }
      // NOTE: No full Graph tree walk here — _flattenVesselViaGraph makes
      // hundreds of API calls per vessel and would cause a request loop.

    });

    if ((result.added && result.added.length > 0) || (result.deleted && result.deleted.length > 0)) {
      // Graph is the live document source. Do not reload the backend flat tree
      // after a delta event: it can contain stale folder IDs and overwrite the
      // just-refreshed SharePoint files. The updated delta map is used for
      // subsequent navigation and uploads.
      if (this._deltaReloadTimer) {
        clearTimeout(this._deltaReloadTimer);
        this._deltaReloadTimer = null;
      }
    }
  }

  // ── HTTP Helpers ──────────────────────────────────────────────────────────

  public _headers(): Record<string, string> {
    const h: Record<string, string> = { 'Content-Type': 'application/json' };
    const sid = this.props.sessionId;
    if (sid) { h['Authorization'] = `Bearer ${sid}`; h['X-Session-ID'] = sid; }
    if (this.props.userEmail) h['X-User-Email'] = this.props.userEmail;
    if (this.props.graphAccessToken) h['X-Graph-Access-Token'] = this.props.graphAccessToken;
    if (this.props.spAccessToken)    h['X-SP-Access-Token']    = this.props.spAccessToken;
    return h;
  }

  /** Records a deletion that already happened client-side (a folder/file
   * deleted directly against Graph rather than through this backend's own
   * delete_folder/delete_file) so "Deleted By" and the optional reason are
   * captured immediately, feeding both the Recycle Bin and the live
   * deletion popup. Best-effort — never blocks or surfaces an error to the
   * user; the native-SPO reconciliation poll backfills anything this misses. */
  public _logDeletion(params: {
    item_type: 'vessel' | 'folder' | 'file';
    drive_item_id?: string | null;
    name: string;
    original_path?: string | null;
    site_name?: string | null;
    site_key?: string | null;
    reason?: string | null;
  }): void {
    fetch(`${this._base()}/api/recycle-bin/log-deletion`, {
      method: 'POST',
      headers: this._headers(),
      body: JSON.stringify(params),
    }).catch(() => undefined);
  }

  private _headersForLocalFallback(): Record<string, string> {
    const headers = this._headers();
    delete headers.Authorization;
    delete headers['X-Session-ID'];
    const localSession = `mock-local-${(this.props.userEmail || 'guest').toLowerCase().replace(/[^a-z0-9]+/g, '-')}`;
    headers['X-Session-ID'] = localSession;
    headers.Authorization = `Bearer ${localSession}`;
    return headers;
  }

  // Upload headers – DO NOT set Content-Type; the browser must set it with the multipart boundary
  public _uploadHeaders(): Record<string, string> {
    const h: Record<string, string> = {};
    const sid = this.props.sessionId;
    if (sid) { h['Authorization'] = `Bearer ${sid}`; h['X-Session-ID'] = sid; }
    if (this.props.userEmail) { h['X-User-Email'] = this.props.userEmail; }
    if (this.props.graphAccessToken) h['X-Graph-Access-Token'] = this.props.graphAccessToken;
    if (this.props.spAccessToken)    h['X-SP-Access-Token']    = this.props.spAccessToken;
    return h;
  }

  // No call through _fetchJson had a client-side timeout: if a backend/Graph
  // request never came back (stuck behind SharePoint Online throttling, a
  // stalled Graph call, a site with a genuine permissions/resolution
  // problem, ...), the returned promise never settled — and every caller's
  // own "loading" state, all the way up to what the Documents module shows,
  // stayed true forever. That's the single root cause behind "switching
  // sites / opening Documents loads for a long time and then hangs": it
  // isn't one call site, it's every call going through here having no
  // ceiling on how long it will wait.
  public static readonly _DEFAULT_FETCH_TIMEOUT_MS = 25000;

  public _fetchJson(url: string, optionsOrSignal?: RequestInit | AbortSignal): Promise<any> {
    const callerProvidedSignal: AbortSignal | undefined = optionsOrSignal && 'aborted' in optionsOrSignal
      ? (optionsOrSignal as AbortSignal)
      : (optionsOrSignal as RequestInit | undefined)?.signal as AbortSignal | undefined;

    // Own an AbortController with a timeout whenever the caller didn't
    // already supply a signal, so a hung request is actually cancelled
    // (freeing its connection-pool slot, same reasoning as
    // _cancelSiteSubtreePrefetch elsewhere in this file) instead of left
    // open in the background forever. A caller that passed its own signal
    // keeps full ownership/control of that signal, unchanged from before.
    let timeoutController: AbortController | null = null;
    let timeoutHandle: number | null = null;
    let effectiveSignal = callerProvidedSignal;
    if (!effectiveSignal) {
      timeoutController = new AbortController();
      effectiveSignal = timeoutController.signal;
      timeoutHandle = window.setTimeout(() => timeoutController!.abort(), VesselEmail._DEFAULT_FETCH_TIMEOUT_MS);
    }

    const opts: RequestInit = optionsOrSignal && 'aborted' in optionsOrSignal
      ? { signal: effectiveSignal, headers: this._headers(), cache: 'no-store' }
      : {
        cache: 'no-store',
        ...(optionsOrSignal as RequestInit || {}),
        signal: effectiveSignal,
        headers: {
          ...this._headers(),
          ...((optionsOrSignal as RequestInit)?.headers || {}),
        },
      };

    const clearOwnTimeout = (): void => {
      if (timeoutHandle !== null) window.clearTimeout(timeoutHandle);
    };

    return fetch(url, opts)
      .then(async r => {
        clearOwnTimeout();
        if (r.status === 401) {
          this.setState({ sessionExpired: true, authPage: 'login' });
          throw new Error('SESSION_EXPIRED');
        }
        // 403 = authenticated but not permitted (e.g. not an admin / no site
        // permission). Re-authenticating cannot fix it, and treating it as an
        // expired session caused an endless sign-in reload loop.
        if (r.status === 403) {
          throw new Error('PERMISSION_DENIED_403');
        }
        if (r.status >= 500) {
          // The backend now returns a JSON {detail: "..."} body describing
          // what actually failed (a Graph error, an unexpected exception,
          // ...) on endpoints like /children — this used to be discarded
          // entirely, so every 500 showed up here as a bare "SERVER_ERROR_500"
          // with nothing to act on, in both the console and DocumentsPage's
          // "Couldn't load this folder" retry card. Include it when present.
          let detail = '';
          try {
            const body = await r.json();
            detail = (body && (body.detail || body.message)) || '';
          } catch {
            // Response wasn't JSON (e.g. a proxy/HTML error page) — fall
            // back to the plain status-only message below.
          }
          throw new Error(detail ? `SERVER_ERROR_${r.status}: ${detail}` : `SERVER_ERROR_${r.status}`);
        }
        if (r.status === 429) {
          // The backend hit Microsoft Graph's per-app request quota
          // (SharePoint Embedded `activityLimitReached`) and set a
          // Retry-After header for how long that cooldown lasts (can be a
          // few minutes, not a few seconds). Surface it distinctly so
          // callers like _loadSiteFolderChildren can stop issuing new
          // requests for that whole window instead of the UI's normal
          // "Retry" button immediately generating another one.
          const retryAfterSec = Number(r.headers.get('Retry-After')) || 60;
          throw new Error(`THROTTLED_429:${retryAfterSec}`);
        }
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.json();
      })
      .catch(async err => {
        clearOwnTimeout();
        if (opts.signal?.aborted || err?.message === 'SESSION_EXPIRED' || err?.message === 'PERMISSION_DENIED_403') throw err;
        // A Graph quota throttle is an application-level condition (the
        // backend told Graph "too much"), not a network/connectivity
        // failure — falling back to a different backend host wouldn't fix
        // it and would just muddy which backend is "active". Let it
        // propagate straight to the caller (see _loadSiteFolderChildren).
        if (typeof err?.message === 'string' && err.message.indexOf('THROTTLED_429:') === 0) throw err;
        // The server answered with an error status (404, 500, ...): it is up,
        // so a local backend can't help. Report the real error instead of
        // switching hosts — one failing endpoint used to flip the whole app
        // to 127.0.0.1 for the rest of the session.
        const answered = typeof err?.message === 'string'
          && (err.message.indexOf('SERVER_ERROR_') === 0 || err.message.indexOf('HTTP ') === 0);
        if (!answered && url.includes('nk-dms-dev.sg-nissenkaiun.com')) {
          const fallbackUrl = url.replace('https://nk-dms-dev.sg-nissenkaiun.com', 'http://127.0.0.1:8000');
          try {
            const localOpts: RequestInit = {
              ...opts,
              headers: this._headersForLocalFallback(),
            };
            const r = await fetch(fallbackUrl, localOpts);
            // Only stick to the local backend once it has actually answered.
            VesselEmail._remoteServerDown = true;
            console.warn(`[VesselDMS] Remote API unreachable (${err?.message}) — switched to local backend: ${fallbackUrl}`);
            if (r.status === 401) {
              this.setState({ sessionExpired: true, authPage: 'login' });
              throw new Error('SESSION_EXPIRED');
            }
            if (r.status === 403) throw new Error('PERMISSION_DENIED_403');
            if (!r.ok) throw new Error(`HTTP ${r.status}`);
            return await r.json();
          } catch (localErr: any) {
            if (localErr?.message === 'SESSION_EXPIRED' || localErr?.message === 'PERMISSION_DENIED_403') throw localErr;
            console.warn('[VesselDMS] Local fallback fetch failed:', localErr);
            return null;
          }
        }
        throw err;
      });
  }

  public static _remoteServerDown: boolean = false;

  public _handleSignOut = (): void => {
    const base = this._base();
    const sid = this.props.sessionId || '';
    const email = this.props.userEmail || '';
    fetch(`${base}/api/auth/logout`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, session_id: sid }),
    }).catch(() => undefined);
    this.setState({ authPage: 'logout', sessionExpired: false });
  };

  public _base(): string {
    const url = (this.props.apiBaseUrl || '').replace(/\/$/, '');
    if (VesselEmail._remoteServerDown && (!url || url.includes('nk-dms-dev.sg-nissenkaiun.com'))) {
      return 'http://127.0.0.1:8000';
    }
    return url || 'https://nk-dms-dev.sg-nissenkaiun.com';
  }

  public async _loadDocumentSites(): Promise<void> {
    if (this._documentSitesLoading) return;
    this._documentSitesLoading = true;
    try {
      let base = this._base();
      let data: any = null;
      try {
        data = await this._fetchJson(`${base}/api/documents/sites`);
      } catch {
        if (!base.includes('localhost')) {
          data = await this._fetchJson('http://127.0.0.1:8000/api/documents/sites').catch(() => null);
        }
      }
      if (!data || !Array.isArray(data.sites) || data.sites.length === 0) {
        try {
          const cfg = await this._fetchJson(`${base}/api/config/available-sites`).catch(() => null);
          if (Array.isArray(cfg?.sites) && cfg.sites.length > 0) {
            data = {
              sites: cfg.sites.map((s: any) => ({
                site_key: s.site_key || s.name,
                sp_site_name: s.sp_site_name || s.display_name || s.name,
                site_id: s.site_id || '',
                drive_id: s.drive_id || '',
                web_url: s.web_url || '',
                default_library_name: s.default_library_name || ((s.name || '').toLowerCase().includes('nks') ? 'Shared Documents' : 'Documents'),
                is_primary: true,
              })),
              active_site: cfg.current_site,
            };
          }
        } catch { /* ignore */ }
      }
      let sites = Array.isArray(data?.sites) ? data.sites : [];
      const active = data?.active_site || sites[0]?.site_key || null;
      const normalizeSiteKey = (value: string | null | undefined): string =>
        String(value || '').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
      const resolveDefaultSiteKey = (requested: string | null, available: Array<{ site_key?: string; sp_site_name?: string }>): string => {
        if (!requested) return available[0]?.site_key || 'all';
        const requestedKey = normalizeSiteKey(requested);
        const exact = available.find(site => normalizeSiteKey(site.site_key) === requestedKey);
        if (exact) return exact.site_key || requested;
        const aliasMatch = available.find(site => {
          const siteKey = normalizeSiteKey(site.site_key);
          const siteName = normalizeSiteKey(site.sp_site_name);
          if (!siteKey && !siteName) return false;
          if (requestedKey === 'local' && (siteKey.includes('dev') || siteKey.includes('communication') || siteName.includes('dev') || siteName.includes('communication'))) return true;
          if (requestedKey === 'dev' && (siteKey.includes('local') || siteKey.includes('docman') || siteName.includes('local') || siteName.includes('docman'))) return true;
          if (requestedKey.includes('communication') && (siteKey.includes('dev') || siteName.includes('dev'))) return true;
          return siteKey === requestedKey || siteName === requestedKey;
        });
        return aliasMatch?.site_key || available[0]?.site_key || 'all';
      };
      const initialSite = sites.find((site: any) => site.site_key === active) || sites[0];
      const defaultSiteFilter = resolveDefaultSiteKey(active, sites);
      const defaultSiteTarget = defaultSiteFilter === 'all' ? (initialSite?.site_key || null) : defaultSiteFilter;
      const wantsInitialSiteRoot = Boolean(
        initialSite?.site_id && initialSite?.drive_id &&
        this.state.folderPathStack.length === 0 &&
        this.state.docScopeType === 'vessels'
      );
      // site_configurations.drive_id can go stale (site re-provisioned, drive
      // recreated, DB entry never updated, etc.). The manual Sites > site >
      // drive click-through always re-resolves the drive live via Graph
      // (_getOrLoadSiteDrives / _loadSiteDrives), so it recovers even when the
      // stored id is wrong. This auto-open-on-load shortcut used to trust
      // initialSite.drive_id directly and jump straight to the folder-root
      // view — when the stored id was stale, every folder request beneath it
      // came back empty/erroring (silently, since _getOrLoadSiteFolderChildren
      // swallows the failure), which looked exactly like "no files here" to
      // the user, but only on this shortcut path — not when navigating there
      // by hand. Verify against Graph's live drive list first.
      //
      // This verification used to run only when wantsInitialSiteRoot was true
      // (docScopeType === 'vessels'). But `initialSite.drive_id` — patched
      // into `documentSites` below — is also what DocumentsPage's level-0
      // view (`activeLiveSite.drive_id`, DocumentsPage.tsx ~L291) uses for
      // its very first root `/children` fetch when the user lands directly
      // on Documents/Shared Documents/Sites scope instead of Vessels. A stale
      // id there made that first fetch 404/403 silently, showing "0 main
      // folders & libraries" with an empty grid until the user manually
      // reselected a site (which re-resolves the drive via
      // _switchDocumentSite). Running this for every initialSite, regardless
      // of scope, fixes it at the source instead of only for the auto-open
      // shortcut.
      let resolvedInitialDrive: { id: string; name: string } | null = initialSite?.drive_id
        ? { id: initialSite.drive_id, name: initialSite.default_library_name || 'Documents' }
        : null;
      let openInitialSiteRoot = wantsInitialSiteRoot;
      if (initialSite?.site_id) {
        try {
          const liveDrives = await this._loadSiteDrives(initialSite.site_id);
          if (liveDrives.length > 0) {
            this._siteDrivesCache.set(initialSite.site_id, liveDrives);
            const preferredDrive = liveDrives.find(d => /^(documents|shared documents)$/i.test((d.name || '').trim()));
            const configuredDrive = liveDrives.find(d => d.id === initialSite.drive_id);
            const matchedDrive = preferredDrive || configuredDrive || liveDrives[0];
            resolvedInitialDrive = { id: matchedDrive.id, name: matchedDrive.name || resolvedInitialDrive?.name || 'Documents' };
            // Patch the verified drive back into `initialSite`/`sites` so
            // every consumer of `documentSites` set below (the site
            // dropdown, DocumentsPage's `activeLiveSite`) sees the live id,
            // not the possibly-stale one from `/api/documents/sites`.
            initialSite.drive_id = resolvedInitialDrive.id;
            initialSite.default_library_name = resolvedInitialDrive.name;
            sites = sites.map((site: any) => site.site_key === initialSite.site_key
              ? { ...site, drive_id: resolvedInitialDrive!.id, default_library_name: resolvedInitialDrive!.name }
              : site);
          } else if (wantsInitialSiteRoot) {
            // Could not verify any live drive for this site — don't auto-drop
            // the user into a folder view that will just render as empty.
            // Land on the Sites list instead; it resolves drives itself.
            openInitialSiteRoot = false;
          }
        } catch {
          // Network/Graph error while verifying — fall back to the configured
          // id rather than blocking the auto-open entirely.
        }
      }
      const initialSiteRoot = openInitialSiteRoot ? [
        { id: 'sites_root', name: 'SharePoint Sites' },
        { id: `site:${initialSite.site_id}`, name: initialSite.sp_site_name || initialSite.site_key },
        { id: `drive:${resolvedInitialDrive!.id}`, name: resolvedInitialDrive!.name },
      ] : this.state.folderPathStack;
      if (openInitialSiteRoot) {
        this.setState({
          documentSites: sites,
          activeDocumentSite: active,
          vesselSiteFilter: this.state.vesselSiteFilter === 'all' || !this.state.vesselSiteFilter ? (defaultSiteTarget || defaultSiteFilter) : this.state.vesselSiteFilter,
          docScopeType: 'sites',
          docViewMode: 'folder',
          folderPathStack: initialSiteRoot,
          docMainFolder: 'SharePoint Sites',
          folderNavHistory: [{ folderPathStack: initialSiteRoot, docMainFolder: 'SharePoint Sites' }],
          folderNavIndex: 0,
        });
      } else {
        // A folder navigation restored from sessionStorage after a refresh
        // (see _readPersistedFolderNav) belongs to whichever site the user was
        // browsing — but a refresh signs in with a NEW session, whose
        // server-side active site is back to the default. Using `active`
        // here made the "SharePoint site" dropdown say e.g. "Communication
        // Site" while the breadcrumb/folders were still NissenKaiunExternal.
        // The restored breadcrumb wins: the dropdown follows it, and the new
        // session is pointed at that site so uploads/actions go to the site
        // being shown.
        const restoredSiteKey = this._siteKeyForFolderStack(this.state.folderPathStack, sites);
        const nextActiveSite = restoredSiteKey || active;
        this.setState({
          documentSites: sites,
          activeDocumentSite: nextActiveSite,
          vesselSiteFilter: this.state.vesselSiteFilter === 'all' || !this.state.vesselSiteFilter ? (defaultSiteTarget || defaultSiteFilter) : this.state.vesselSiteFilter,
        });
        if (restoredSiteKey && restoredSiteKey !== active) {
          void this._syncSessionActiveSite(restoredSiteKey).catch(error => {
            // Couldn't align the session: fall back to the session's own
            // site and its library root, so the dropdown and the folders
            // shown still agree.
            console.warn('[VesselDMS] Could not restore active site after refresh:', error);
            if (active && sites.some((site: any) => site.site_key === active)) {
              void this._switchDocumentSite(active).catch(() => undefined);
            }
          });
        }
      }
      // Fetch dynamic aliases alongside sites
      void this._loadDocumentAliases().catch(() => undefined);
      // state.folderPathStack/docScopeType above are now either a freshly
      // live-verified root (openInitialSiteRoot) or a sessionStorage restore
      // that's just been re-synced against the live session (the `else`
      // branch above) — either way, safe for _goToView('list') to treat as
      // real, already-loaded live data on the user's first Documents click.
      this._initialDocumentsRootReady = true;
    } catch (error) {
      console.warn('[VesselDMS] configured Documents sites unavailable:', error);
    } finally {
      this._documentSitesLoading = false;
    }
  }

  public async _loadDocumentAliases(): Promise<void> {
    try {
      const base = this._base();
      let data = await this._fetchJson(`${base}/api/documents/aliases`).catch(() => null);
      if (!data && !base.includes('localhost')) {
        data = await this._fetchJson('http://127.0.0.1:8000/api/documents/aliases').catch(() => null);
      }
      if (data) {
        this.setState({
          documentDepartmentAliases: data.departments || {},
          documentVesselAliases: data.vessels || {},
        });
      }
    } catch { /* non-fatal */ }
  }

  private _allVesselNamesForSearchLoading = false;
  /** Fetches every vessel name tenant-wide (no site_key), once, purely so
   * DocumentsPage's search-clause classifier can recognize a query clause
   * as naming a vessel even when browsing a SharePoint site/library whose
   * `vessels` (the Vessels-module table — scoped by vesselSiteFilter, and
   * deliberately left un-refetched on a Documents site switch, see
   * _switchDocumentSite's comment) doesn't include that vessel. Never
   * rendered anywhere itself, so being tenant-wide/unscoped is safe. */
  public async _loadAllVesselNamesForSearch(): Promise<void> {
    if (this._allVesselNamesForSearchLoading || this.state.allVesselNamesForSearch.length > 0) return;
    this._allVesselNamesForSearchLoading = true;
    try {
      const base = this._base();
      const data = await this._fetchJson(`${base}/api/vessels`).catch(() => null);
      if (Array.isArray(data)) {
        const names = data.map((v: any) => String(v?.name || '').trim()).filter(Boolean);
        if (names.length > 0) {
          this.setState({ allVesselNamesForSearch: names });
        }
      }
    } catch { /* non-fatal — search just falls back to whatever `vessels` already has */ }
    finally {
      this._allVesselNamesForSearchLoading = false;
    }
  }

  public _tenantSitesLoading = false;
  public async _loadTenantSites(limit: number = 50): Promise<Array<{ id: string; name: string; display_name: string; web_url?: string }>> {
    if (this.state.tenantDiscoveredSites && this.state.tenantDiscoveredSites.length > 0) {
      return this.state.tenantDiscoveredSites;
    }
    try {
      this._tenantSitesLoading = true;
      const base = this._base();
      let sResp = await this._fetchJson(`${base}/api/sites?limit=${limit}`).catch(() => null);
      if (!sResp && !base.includes('localhost')) {
        sResp = await this._fetchJson(`http://127.0.0.1:8000/api/sites?limit=${limit}`).catch(() => null);
      }
      const rawSites: any[] = Array.isArray(sResp?.sites) ? sResp.sites : [];
      // Deduplicate: Exclude any site that is already in documentSites (configured first-class libraries)
      const configuredIds = new Set<string>();
      (this.state.documentSites || []).forEach(cs => {
        if (cs.site_id) configuredIds.add(cs.site_id.toLowerCase());
        if (cs.site_key) configuredIds.add(cs.site_key.toLowerCase());
        if (cs.sp_site_name) configuredIds.add(cs.sp_site_name.toLowerCase());
      });
      const filtered = rawSites.filter(s => {
        const id = (s.id || '').toLowerCase();
        const name = (s.name || '').toLowerCase();
        const disp = (s.display_name || '').toLowerCase();
        return !configuredIds.has(id) && !configuredIds.has(name) && !configuredIds.has(disp);
      }).map(s => ({
        id: s.id,
        name: s.name || s.id,
        display_name: s.display_name || s.name || s.id,
        web_url: s.web_url,
      }));
      this.setState({ tenantDiscoveredSites: filtered });
      return filtered;
    } catch {
      return [];
    } finally {
      this._tenantSitesLoading = false;
    }
  }

  public async _loadDocumentLiveTree(siteKey: string, knownSites = this.state.documentSites): Promise<void> {
    if (this._documentLiveTreeLoading.has(siteKey) || this._documentLiveTreeLoaded.has(siteKey)) return;
    const site = knownSites.find(item => item.site_key === siteKey);
    if (!site || !site.drive_id) return;
    this._documentLiveTreeLoading.add(siteKey);
    const controller = new AbortController();
    this._documentLiveTreeAbortControllers.set(siteKey, controller);
    const selectedSite = knownSites.find(item => item.site_key === siteKey);
    const isNksDocMan = /nksdocman/i.test(`${selectedSite?.site_key || ''} ${selectedSite?.sp_site_name || ''}`);
    this.setState({
      activeDocumentSite: siteKey,
      documentLiveFoldersLoading: true,
    });
    if (isNksDocMan) {
      this.setState({
        docViewMode: 'folder',
        docScopeType: 'sites',
        folderPathStack: [],
        folderNavHistory: [],
        folderNavIndex: -1,
        docMainFolder: null,
        vesselFilter: 'all',
      });
    }
    try {
      const url = `${this._base()}/api/sites/${encodeURIComponent(site.site_id)}/drives/${encodeURIComponent(site.drive_id)}/folders/root/recursive`;
      const timeout = new Promise<never>((_, reject) => {
        window.setTimeout(() => reject(new Error('LIVE_TREE_TIMEOUT')), 20000);
      });
      const data = await Promise.race([this._fetchJson(url, controller.signal), timeout]);
      if (this.state.activeDocumentSite !== siteKey) return;
      const liveItems = Array.isArray(data?.folders) ? data.folders : [];
      const liveRows: FlatRow[] = [];
      const siteLabel = site.sp_site_name || site.site_key;
      const libraryName = site.default_library_name || 'Shared Documents';
      const knownVessels = this.state.vessels || [];
      liveItems.forEach((item: any, index: number) => {
        const parts = String(item.path || item.name || '').split('/').filter(Boolean);
        if (parts.length === 0 || !item.id) return;
        const folderParts = item.is_folder ? parts : parts.slice(0, -1);
        const folderPath = folderParts.length > 0 ? folderParts : [siteLabel];

        const tags = item.tags || {};
        const tagVessel = tags.vessel && tags.vessel !== 'To Be Classified' && tags.vessel !== 'Unknown' && tags.vessel !== 'Not Listed' ? tags.vessel : null;
        const tagDept = tags.department || tags.mainFolder || null;
        const tagGroup = tags.group || null;
        const tagCat = tags.category || null;
        const tagSubCat = tags.sub_category || tags.subCategory || null;
        const tagDocSec = tags.document_section || tags.documentSection || null;

        const matchedVessel = knownVessels.find(vessel =>
          folderPath.some(part => part.trim().toLowerCase() === vessel.name.trim().toLowerCase())
        );
        const vesselIndex = matchedVessel
          ? folderPath.findIndex(part => part.trim().toLowerCase() === matchedVessel.name.trim().toLowerCase())
          : -1;
        const vesselName = tagVessel || matchedVessel?.name || 'Vessel not assigned';
        const folderLevels = vesselIndex >= 0 ? folderPath.slice(vesselIndex + 1) : folderPath.slice(1);
        const category = tagCat || (folderLevels[0] && folderLevels[0] !== siteLabel ? folderLevels[0] : 'Category not assigned');
        const subCategory = tagSubCat || (folderLevels.length > 1 ? folderLevels[folderLevels.length - 1] : 'Sub-category not assigned');
        const mainFolder = tagDept || 'Technical & Crewing';
        const group = tagGroup || 'Group not assigned';
        const documentSection = tagDocSec || 'Document section not assigned';
        const subFolderPath = ['SharePoint Sites', siteLabel, libraryName, ...folderPath.filter((part, partIndex) => partIndex !== 0 || part.toLowerCase() !== libraryName.toLowerCase())].join(' > ');
        const folderId = item.is_folder ? item.id : item.parent_id;
        if (!folderId) return;
        liveRows.push({
          srNo: `LIVE-${index + 1}`,
          vesselName,
          group: mainFolder,
          documentSection,
          groupTag: group,
          category,
          subCategory,
          subFolderPath,
          fileName: item.is_folder ? null : (item.name || null),
          fileId: item.is_folder ? null : item.id,
          fileUploadedAt: item.created_date_time ? Date.parse(item.created_date_time) : (item.last_modified_date_time ? Date.parse(item.last_modified_date_time) : undefined),
          fileSize: typeof item.size === 'number' ? `${(item.size / 1024).toFixed(1)} KB` : undefined,
          canUpload: false,
          groupKey: `live:${siteKey}:${folderId}`,
          uploadFolderId: folderId,
          monthDriven: false,
          tags,
        });
      });
      await new Promise<void>(resolve => {
        this.setState(previous => ({
          activeDocumentSite: siteKey,
          documentLiveFolders: liveItems,
          rows: [
            ...previous.rows.filter(row => !row.groupKey.startsWith('live:')),
            ...liveRows,
          ],
          documentLiveFoldersLoading: false,
        }), resolve);
      });
    } catch (error: any) {
      // Non-fatal: recursive tree is optional, folder browsing uses direct children
      if (error?.name !== 'AbortError' && error?.message !== 'REQUEST_ABORTED') {
        console.warn('[VesselDMS] live folder scan skipped:', error);
      }
      this._documentLiveTreeLoaded.delete(siteKey); // allow retry
    } finally {
      this._documentLiveTreeAbortControllers.delete(siteKey);
      this._documentLiveTreeLoading.delete(siteKey);
      this._documentLiveTreeLoaded.add(siteKey);
      if (this.state.activeDocumentSite === siteKey) {
        this.setState({ documentLiveFoldersLoading: false });
      }
    }
  }

  /**
   * A site's live folder tree (`documentLiveFolders`) is loaded once per
   * session and then marked done in `_documentLiveTreeLoaded` — see the
   * early-return guard at the top of `_loadDocumentLiveTree`. That guard
   * means a folder created after the initial load (e.g. a new vessel via
   * the "Create a new vessel" dialog) never appeared in the Documents
   * module for that site again until a full page reload, because nothing
   * ever cleared it. Call this right after an operation that adds/removes
   * a root-level folder in a site's drive (vessel create/delete) so the
   * next read re-fetches instead of serving the stale in-memory snapshot.
   */
  public _invalidateDocumentLiveTree(siteKey: string | undefined): void {
    if (!siteKey) return;
    this._documentLiveTreeLoaded.delete(siteKey);
    this._documentLiveTreeLoading.delete(siteKey);
    const site = (this.state.documentSites || []).find(s => s.site_key === siteKey);
    if (site && site.site_id && site.drive_id) {
      const prefix = `${site.site_id}::${site.drive_id}::`;
      Array.from(this._siteFolderItemsCache.keys())
        .filter(key => key.startsWith(prefix))
        .forEach(key => this._siteFolderItemsCache.delete(key));
      Array.from(this._siteFolderTreeCache.keys())
        .filter(key => key.startsWith(prefix))
        .forEach(key => this._siteFolderTreeCache.delete(key));
    }
    // If the user is already looking at this site's Documents view, refresh
    // it right away instead of waiting for them to navigate away and back.
    if (this.state.activeDocumentSite === siteKey) {
      void this._loadDocumentLiveTree(siteKey);
    }
  }

  public _cancelDocumentLiveTree(): void {
    this._documentLiveTreeAbortControllers.forEach(controller => controller.abort());
    this._documentLiveTreeAbortControllers.clear();
    this._documentLiveTreeLoading.clear();
    if (!this._isUnmounted && this.state.documentLiveFoldersLoading) {
      this.setState({ documentLiveFoldersLoading: false });
    }
  }

  public async _loadSiteDrives(siteId: string, signal?: AbortSignal): Promise<Array<{ id: string; name: string; web_url?: string; is_system?: boolean; item_count?: number | null }>> {
    try {
      const data = await this._fetchJson(`${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives`, signal);
      return Array.isArray(data?.drives) ? data.drives : [];
    } catch (err) {
      console.warn('[VesselDMS] _loadSiteDrives error:', err);
      return [];
    }
  }

  public async _loadTermStoreVessels(siteId: string): Promise<string[]> {
    try {
      const data = await this._fetchJson(`${this._base()}/api/sites/${encodeURIComponent(siteId)}/term-store-vessels`);
      return Array.isArray(data?.vessels) ? data.vessels : [];
    } catch (err) {
      console.warn('[VesselDMS] _loadTermStoreVessels error:', err);
      return [];
    }
  }

  // Set whenever the backend reports Microsoft Graph's per-app request
  // quota is exhausted (activityLimitReached) — see _fetchJson's
  // THROTTLED_429 handling. While Date.now() is before this timestamp,
  // _loadSiteFolderChildren fails fast locally instead of issuing another
  // request. Without this, rapid folder navigation or vessel-filter changes
  // kept firing fresh /children requests during the cooldown (each click a
  // new request, each request its own retry attempts on the backend), which
  // is what was re-triggering the throttle instead of letting it clear.
  public _graphThrottledUntil: number = 0;

  // Global cap on simultaneous /children requests. Render-time walks (e.g. the
  // Documents Sub-folder dropdown) and site switches used to fire dozens at
  // once, tripping Graph 429 throttling. Extra requests wait their turn here.
  private _childrenInFlight: number = 0;
  private _childrenWaiters: Array<() => void> = [];
  private async _withChildrenSlot<T>(signal: AbortSignal | undefined, fn: () => Promise<T>): Promise<T> {
    while (this._childrenInFlight >= 4) {
      // eslint-disable-next-line no-await-in-loop
      await new Promise<void>(resolve => { this._childrenWaiters.push(resolve); });
      if (signal?.aborted) {
        const abortErr: any = new Error('REQUEST_ABORTED');
        abortErr.name = 'AbortError';
        throw abortErr;
      }
    }
    this._childrenInFlight++;
    try {
      return await fn();
    } finally {
      this._childrenInFlight--;
      const next = this._childrenWaiters.shift();
      if (next) next();
    }
  }

  public async _loadSiteFolderChildren(
    siteId: string,
    driveId: string,
    folderId: string = 'root',
    signal?: AbortSignal,
  ): Promise<{ items: any[]; summaryCounts?: any; parentPath?: string; folderId?: string; error?: boolean; throttled?: boolean }> {
    if (Date.now() < this._graphThrottledUntil) {
      return { items: [], summaryCounts: null, parentPath: '', folderId, error: true, throttled: true };
    }
    try {
      const encodedFolderRef = folderId
        .split('/')
        .map(part => encodeURIComponent(part))
        .join('/');
      const data = await this._withChildrenSlot(signal, async () => {
        // Throttling may have started while this request was queued.
        if (Date.now() < this._graphThrottledUntil) throw new Error(`THROTTLED_429:${Math.max(1, Math.ceil((this._graphThrottledUntil - Date.now()) / 1000))}`);
        return this._fetchJson(
          `${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/folders/${encodedFolderRef}/children`,
          signal
        );
      });
      return {
        items: Array.isArray(data?.items) ? data.items : [],
        summaryCounts: data?.summary_counts || null,
        parentPath: data?.parent_path || '',
        folderId: data?.folder_id || folderId,
      };
    } catch (err: any) {
      if (typeof err?.message === 'string' && err.message.indexOf('THROTTLED_429:') === 0) {
        const retryAfterSec = Number(err.message.split(':')[1]) || 60;
        const until = Date.now() + retryAfterSec * 1000;
        // Requests queued behind the slot rethrow the same window locally —
        // log once per window, not once per queued folder.
        if (until > this._graphThrottledUntil + 2000) {
          console.warn(`[VesselDMS] _loadSiteFolderChildren throttled by Graph quota — pausing folder requests for ${retryAfterSec}s.`);
        }
        this._graphThrottledUntil = Math.max(this._graphThrottledUntil, until);
        return { items: [], summaryCounts: null, parentPath: '', folderId, error: true, throttled: true };
      }
      // Distinguish "the request failed" from "the folder is genuinely
      // empty" — callers used to treat both the same, which showed a
      // misleading "This folder is empty" for e.g. a stale/wrong drive id, a
      // 404, or a transient network error.
      console.warn('[VesselDMS] _loadSiteFolderChildren error:', err);
      return { items: [], summaryCounts: null, parentPath: '', folderId, error: true };
    }
  }

  public _siteDrivesCache: Map<string, Array<{ id: string; name: string; web_url?: string; is_system?: boolean; item_count?: number | null }>> = new Map();
  // Site ids for which the "Libraries in <site>" grid (DocumentsPage.tsx,
  // stackLevel 2) should also show SharePoint's own auto-provisioned system
  // libraries (Site Assets, Style Library, Form Templates, ...), which are
  // hidden by default there for the same reason they're hidden in the Sites
  // module — they're never where a user's actual files live. Toggled by the
  // "Show N more" control in that grid; a plain Set + forceUpdate() rather
  // than state, matching how _siteDrivesCache itself is handled.
  public _systemLibrariesRevealed: Set<string> = new Set();
  public _siteFolderItemsCache: Map<string, { items: any[]; loading?: boolean; parentPath?: string; error?: boolean; throttled?: boolean; retryAt?: number; errorAt?: number; _promise?: Promise<any[]> }> = new Map();
  // Consecutive non-throttle failures per folder listing (key = cache key).
  // A failed listing retries by itself with a growing delay (see
  // _folderListingRetryDelayMs) instead of sitting on "Couldn't load this
  // folder" until the user presses Retry.
  public _folderListingFailures: Map<string, number> = new Map();
  public _folderListingRetryDelayMs(key: string, failed: boolean): number {
    if (!failed) { this._folderListingFailures.delete(key); return 0; }
    const n = (this._folderListingFailures.get(key) || 0) + 1;
    this._folderListingFailures.set(key, n);
    // 3 s, 6 s, 12 s, 24 s, then leave the manual Retry.
    return n <= 4 ? 3000 * Math.pow(2, n - 1) : 0;
  }

  /** Folder-only tree for a folder, fetched with ONE backend call
   * (GET .../folders/{id}/folder-tree) instead of one /children call per
   * folder. Feeds the Documents "All sub-folders" dropdown. Flat, parent-linked
   * rows; `truncated` means the server hit a cap or a branch failed. */
  public _siteFolderTreeCache: Map<string, {
    folders: Array<{ id: string; name: string; parent_id: string; path: string; depth: number }>;
    truncated: boolean; loading: boolean; error: boolean; ts: number;
  }> = new Map();

  /** Folder-tree keys with a poll timer pending (see _fetchSiteFolderTree). */
  private _siteFolderTreePolling: Set<string> = new Set();

  public _getOrLoadSiteFolderTree(siteId: string, driveId: string, folderId: string, maxDepth?: number, maxFolders?: number): {
    folders: Array<{ id: string; name: string; parent_id: string; path: string; depth: number }>;
    truncated: boolean; loading: boolean; error: boolean;
  } {
    const empty = { folders: [], truncated: false, loading: false, error: false };
    if (!siteId || !driveId || !folderId) return empty;
    // maxDepth: a shallow tree (e.g. the whole library 3 levels deep, used to
    // find vessel folders for the Vessel filter) — cached separately from the
    // full-depth tree of the same folder.
    // maxFolders: raise the server's folder cap (default 6000) for a tree that
    // must reach the deepest levels, e.g. the whole library's categories.
    const key = `${siteId}::${driveId}::${folderId}${maxDepth ? `::d${maxDepth}` : ''}${maxFolders ? `::f${maxFolders}` : ''}`;
    const hit = this._siteFolderTreeCache.get(key);
    const STALE_MS = 5 * 60 * 1000;
    // Also re-ask when the last answer was an empty "still building" tree and
    // no poll is scheduled any more (the poll chain below gives up after a
    // while) — otherwise e.g. the library-wide Sub-category list stayed empty
    // until the 5-minute staleness window passed.
    const retryOk = !!hit && (hit.error
      ? Date.now() - hit.ts > 30000
      : (hit.truncated && hit.folders.length === 0 && !this._siteFolderTreePolling.has(key) && Date.now() - hit.ts > 20000));
    if (hit && hit.loading) return hit;
    if (hit && !retryOk && Date.now() - hit.ts <= STALE_MS) return hit;
    // Never add load while Graph is throttling; serve whatever we have.
    if (Date.now() < this._graphThrottledUntil) return hit || empty;
    this._fetchSiteFolderTree(key, siteId, driveId, folderId, false, 0, maxDepth, maxFolders);
    return this._siteFolderTreeCache.get(key)!;
  }

  private _fetchSiteFolderTree(key: string, siteId: string, driveId: string, folderId: string, refresh: boolean, attempt: number, maxDepth?: number, maxFolders?: number): void {
    const prev = this._siteFolderTreeCache.get(key);
    // Stale-while-revalidate: keep serving the previous tree while refreshing.
    const base = prev || { folders: [], truncated: false, loading: true, error: false, ts: Date.now() };
    this._siteFolderTreeCache.set(key, { ...base, loading: true });
    const encodedFolderRef = folderId.split('/').map(p => encodeURIComponent(p)).join('/');
    // A cold build of a big library can take longer than the default 25 s
    // fetch timeout (that abort is what logged "folder-tree load failed:
    // AbortError"), so give this one call its own, longer deadline.
    const treeAbort = new AbortController();
    const treeTimer = window.setTimeout(() => treeAbort.abort(), 120000);
    this._fetchJson(
      `${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/folders/${encodedFolderRef}/folder-tree` +
        (() => {
          const qs = [refresh ? 'refresh=true' : '', maxDepth ? `max_depth=${maxDepth}` : '', maxFolders ? `max_folders=${maxFolders}` : ''].filter(Boolean).join('&');
          return qs ? `?${qs}` : '';
        })(),
      treeAbort.signal
    ).then((data: any) => {
      window.clearTimeout(treeTimer);
      const folders = Array.isArray(data?.folders) ? data.folders : [];
      const truncated = !!data?.truncated;
      // Never let a smaller partial (throttled) result replace a bigger one.
      // (A complete delta-index answer is authoritative even when depth-capped,
      // so a folder deleted since the last load does drop out.)
      const partialAnswer = !!data?.building || data?.source !== 'delta';
      const keepOld = truncated && partialAnswer && base.folders.length > folders.length;
      this._siteFolderTreeCache.set(key, keepOld
        ? { ...base, loading: false, error: false, ts: Date.now() }
        : { folders, truncated, loading: false, error: false, ts: Date.now() });
      // Still building on the server (first drive index pass) or a partial
      // tree: poll again at a steady pace — never faster than the server's
      // own retry_after, and never hammering during a quota cooldown.
      // A shallow tree is "truncated" by design (depth cap) — only poll it
      // while the server index is still building.
      const pollPartial = !maxDepth && data?.source !== 'delta';
      // "building" polls cost the server no Graph calls, so keep asking for
      // longer (a cold index build under throttling can take minutes).
      if (truncated && (data?.building || pollPartial) && attempt < (data?.building ? 60 : 15)) {
        const serverWaitMs = Number(data?.retry_after) > 0 ? Number(data.retry_after) * 1000 + 1000 : 0;
        const delayMs = Math.min(60000, Math.max(serverWaitMs, data?.building ? 5000 : 8000));
        this._siteFolderTreePolling.add(key);
        window.setTimeout(() => {
          this._siteFolderTreePolling.delete(key);
          this._fetchSiteFolderTree(key, siteId, driveId, folderId, false, attempt + 1, maxDepth, maxFolders);
        }, delayMs);
      }
    }).catch((err: any) => {
      window.clearTimeout(treeTimer);
      if (typeof err?.message === 'string' && err.message.indexOf('THROTTLED_429:') === 0) {
        const retryAfterSec = Number(err.message.split(':')[1]) || 60;
        this._graphThrottledUntil = Math.max(this._graphThrottledUntil, Date.now() + retryAfterSec * 1000);
      }
      console.warn('[VesselDMS] folder-tree load failed:', err);
      this._siteFolderTreeCache.set(key, { ...base, loading: false, error: true, ts: Date.now() });
      // Retry once the throttle window has passed (not every few seconds).
      const waitMs = Math.max(10000, this._graphThrottledUntil - Date.now() + 1000);
      if (attempt < 6) {
        window.setTimeout(() => this._fetchSiteFolderTree(key, siteId, driveId, folderId, false, attempt + 1, maxDepth, maxFolders), Math.min(waitMs, 90000));
      }
    }).then(() => this._scheduleForceUpdate());
  }

  /**
   * Server-computed folder counts for the Documents folder tiles — the same
   * /subfolder-counts endpoint the Sites page uses. The live recursive tree
   * (documentLiveFolders) is capped at 2000 items server-side, so counts
   * derived from it are partial on large libraries.
   */
  public _folderCountsCache: Map<string, {
    ts: number;
    loading: boolean;
    polls?: number;
    counts: Record<string, { direct_subfolders: number; direct_files: number; total_subfolders: number; total_files: number }>;
    summary: { direct_folders: number; direct_files: number; total_folders: number; total_files: number } | null;
  }> = new Map();

  public _getOrLoadFolderCounts(siteId: string, driveId: string, folderId: string = 'root'): {
    counts: Record<string, { direct_subfolders: number; direct_files: number; total_subfolders: number; total_files: number }>;
    summary: { direct_folders: number; direct_files: number; total_folders: number; total_files: number } | null;
  } | null {
    if (!siteId || !driveId) return null;
    const key = `${siteId}::${driveId}::${folderId}`;
    const hit = this._folderCountsCache.get(key);
    if (hit && (hit.loading || Date.now() - hit.ts < 60000)) {
      return hit.ts > 0 ? hit : null;
    }
    // No Graph throttle guard here: the server answers from its drive index
    // (no Graph calls), and skipping this call during a cooldown made every
    // folder tile fall back to its own /children request afterwards.
    const entry = { ts: hit?.ts || 0, loading: true, polls: hit?.polls || 0, counts: hit?.counts || {}, summary: hit?.summary || null };
    this._folderCountsCache.set(key, entry);
    const url = `${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/folders/${encodeURIComponent(folderId)}/subfolder-counts`;
    this._withChildrenSlot(undefined, () => this._fetchJson(url)).then(data => {
      // The server answers from its drive index; on first use it is still
      // building. Stay "loading" (so tiles don't fall back to one /children
      // call each) and ask again shortly — the poll costs no Graph calls.
      if (data?.building && (entry.polls || 0) < 30) {
        const waitMs = Math.min(30000, Math.max(5000, (Number(data.retry_after) || 0) * 1000));
        this._folderCountsCache.set(key, { ...entry, loading: true, polls: (entry.polls || 0) + 1 });
        window.setTimeout(() => {
          const cur = this._folderCountsCache.get(key);
          if (cur && cur.loading) this._folderCountsCache.set(key, { ...cur, loading: false, ts: 0 });
          this._scheduleForceUpdate();
        }, waitMs);
        return;
      }
      this._folderCountsCache.set(key, { ts: Date.now(), loading: false, polls: 0, counts: data?.counts || {}, summary: data?.summary_counts || null });
      this._scheduleForceUpdate();
    }).catch((err: any) => {
      if (typeof err?.message === 'string' && err.message.indexOf('THROTTLED_429:') === 0) {
        const retryAfterSec = Number(err.message.split(':')[1]) || 60;
        this._graphThrottledUntil = Math.max(this._graphThrottledUntil, Date.now() + retryAfterSec * 1000);
      }
      // Keep any previous counts; retry after the TTL.
      this._folderCountsCache.set(key, { ts: Date.now(), loading: false, counts: entry.counts, summary: entry.summary });
    });
    return entry.ts > 0 ? entry : null;
  }
  /** True while the /subfolder-counts request for this folder is in flight. */
  public _folderCountsLoading(siteId: string, driveId: string, folderId: string = 'root'): boolean {
    const hit = this._folderCountsCache.get(`${siteId}::${driveId}::${folderId}`);
    return !!hit && hit.loading;
  }
  private _siteFolderChildPrefetched: Set<string> = new Set();
  // Folder ids already auto-drilled-through (see _getOrLoadSiteFolderChildren
  // below): a folder whose only content is exactly one sub-folder and no
  // files gets skipped straight into automatically, once, so browsing a
  // vessel folder that only contains e.g. "Drawings and Manuals" doesn't
  // force an extra click through an otherwise-empty single-item listing.
  private _autoDrilledFolderIds: Set<string> = new Set();
  public _termStoreVesselsCache: Map<string, string[]> = new Map();
  // Tracks the AbortController for the currently-running _prefetchSiteSubtree
  // walk, if any. See _cancelSiteSubtreePrefetch below for why this exists:
  // that background walk's own outstanding requests share the browser's
  // small per-origin connection pool with whatever single folder request
  // Folder view needs next, and were starving it (Folder view stuck on
  // "Loading contents of X..." while List view, which never awaits the
  // walk, rendered instantly).
  private _subtreePrefetchAbort: AbortController | null = null;

  // Tracks the in-flight `_switchDocumentSite` call, if any, the same way
  // `_subtreePrefetchAbort` tracks the subtree walk. Repeatedly changing the
  // "SharePoint site" dropdown used to leave every earlier switch's requests
  // (the /api/sites/active POST, the drives lookup, and the new site's root
  // `/children` fetch) running to completion uncancelled — each one holding
  // a slot in the browser's small per-origin connection pool and queuing up
  // behind one another, which is what showed up as a pile-up of "children"
  // calls in the Network tab and a "Loading contents of Documents..." that
  // never finished after clicking through a few sites in a row. `_siteSwitchGeneration`
  // is bumped on every call and captured locally so a switch that's been
  // superseded by a newer one can tell and bail out without applying its
  // (now stale) results or surfacing a spurious "Could not switch site" alert.
  private _siteSwitchAbort: AbortController | null = null;
  private _siteSwitchGeneration: number = 0;

  // Async cache-fill completions (term store vessels, site drives, site
  // folder children, ...) each used to call this.forceUpdate() directly.
  // That was fine one at a time, but a burst of several completing within
  // the same tick — e.g. a recursive SharePoint subtree prefetch (see
  // _prefetchSiteSubtree below) resolving a dozen folders in quick
  // succession — fired a full synchronous re-render of this ~9,000-line
  // root component once per completion. Repeated navigation between Folder
  // view and List view, each triggering its own burst, is what made the
  // app hang. _scheduleForceUpdate coalesces any completions within a
  // short window into a single re-render.
  private _forceUpdateScheduled: boolean = false;
  public _scheduleForceUpdate(): void {
    if (this._forceUpdateScheduled) return;
    this._forceUpdateScheduled = true;
    setTimeout(() => {
      this._forceUpdateScheduled = false;
      this.forceUpdate();
    }, 120);
  }

  /** Promise-based, cache-aware folder-children loader — shares
   * _siteFolderItemsCache with the synchronous _getOrLoadSiteFolderChildren
   * (same key format), but returns a Promise so callers (notably
   * _prefetchSiteSubtree) can await one folder's load before recursing into
   * its children instead of firing every request at once. */
  public _loadAndCacheSiteFolderChildren(siteId: string, driveId: string, folderId: string, signal?: AbortSignal): Promise<any[]> {
    if (!siteId || !driveId) return Promise.resolve([]);
    const key = `${siteId}::${driveId}::${folderId}`;
    const cached = this._siteFolderItemsCache.get(key) as any;
    if (cached && !cached.loading) {
      // A failed listing is not an answer ("the folder has no such child"):
      // once its retry window has passed, load it again instead of handing
      // every caller the empty items of the error entry.
      const retryable = cached.error === true && (cached.retryAt
        ? Date.now() >= cached.retryAt
        : Date.now() - (cached.errorAt || 0) > 3000);
      if (!retryable) return Promise.resolve(cached.items || []);
    }
    if (cached && cached.loading && cached._promise) return cached._promise;

    const promise: Promise<any[]> = this._loadSiteFolderChildren(siteId, driveId, folderId, signal).then(res => {
      const loadedEntry: { items: any[]; loading: boolean; parentPath: string; error: boolean; throttled?: boolean; retryAt?: number; errorAt?: number } = { items: res.items || [], loading: false, parentPath: res.parentPath || '', error: res.error === true };
      if (res.throttled) { loadedEntry.throttled = true; loadedEntry.retryAt = Math.max(this._graphThrottledUntil, Date.now() + 1000); }
      if (loadedEntry.error) loadedEntry.errorAt = Date.now();
      this._folderListingRetryDelayMs(key, loadedEntry.error && !res.throttled);
      this._siteFolderItemsCache.set(key, loadedEntry);
      this._scheduleForceUpdate();
      return loadedEntry.items;
    }).catch(() => {
      // An aborted background prefetch (see _cancelSiteSubtreePrefetch) is
      // not a real load failure — drop the cache entry instead of marking it
      // errored so a later real request for this folder retries cleanly
      // rather than showing a stuck error/empty state.
      if (signal?.aborted) {
        this._siteFolderItemsCache.delete(key);
      } else {
        this._siteFolderItemsCache.set(key, { items: [], loading: false, parentPath: '', error: true });
      }
      this._scheduleForceUpdate();
      return [];
    });
    this._siteFolderItemsCache.set(key, { items: [], loading: true, _promise: promise } as any);
    return promise;
  }

  /** Recursively prefetches a SharePoint site's folder subtree so List View
   * can flatten every file underneath `folderId`, not just the folders the
   * user has already opened one at a time in Folder view. Bounded by both
   * depth and total node count so a very large or deep library can't hang
   * the browser or flood Graph with requests; already-cached folders
   * resolve instantly via _loadAndCacheSiteFolderChildren, so re-entering a
   * folder already walked is cheap.
   *
   * Budget/concurrency were previously 80/6 — every single Folder -> List
   * switch in the Documents/Sites module could burst up to 80 concurrent-ish
   * requests to our own /api/sites/.../children endpoint (6 at a time),
   * each of which can itself retry up to 6 times server-side against Graph
   * on 429/503 (see GraphClient.request). That's what showed up as the
   * "children API is looping" — a rapid flood of children calls in the
   * Network tab — and is also what was tipping some sites into SharePoint
   * Online's own throttling (the Throttle.htm page reported earlier).
   * Halving both trims the peak burst substantially while still covering
   * realistically-sized folder trees; already-cached folders still resolve
   * instantly on repeat navigation regardless of this budget. */
  public async _prefetchSiteSubtree(
    siteId: string,
    driveId: string,
    folderId: string,
    depth: number = 4,
    budget: { remaining: number } = { remaining: 40 },
  ): Promise<void> {
    if (!siteId || !driveId || depth <= 0 || budget.remaining <= 0) return;

    // Cancel any previous walk before starting a new one — otherwise
    // repeated Folder view <-> List view switches pile up multiple
    // overlapping background walks, each holding several of the browser's
    // small per-origin connection slots and starving whatever single folder
    // request Folder view needs next.
    this._subtreePrefetchAbort?.abort();
    const abortController = new AbortController();
    this._subtreePrefetchAbort = abortController;
    const { signal } = abortController;

    // Breadth-first, a few folders at a time with one overall concurrency
    // cap. The previous depth-first walk awaited every folder strictly one
    // after another, so List view waited for up to `budget` round-trips back
    // to back. Lowered from 6 to reduce how many requests land on the
    // backend/Graph at once (see the "children API is looping" note above).
    const CONCURRENCY = 3;
    let level: string[] = [folderId];
    for (let remainingDepth = depth; remainingDepth > 0 && level.length > 0 && budget.remaining > 0; remainingDepth--) {
      if (signal.aborted) return;
      const nextLevel: string[] = [];
      for (let i = 0; i < level.length && budget.remaining > 0; i += CONCURRENCY) {
        if (signal.aborted) return;
        const batch = level.slice(i, i + Math.min(CONCURRENCY, budget.remaining));
        budget.remaining -= batch.length;
        // eslint-disable-next-line no-await-in-loop
        const results = await Promise.all(batch.map(id => this._loadAndCacheSiteFolderChildren(siteId, driveId, id, signal)));
        if (signal.aborted) return;
        results.forEach(items => {
          (items || []).forEach(item => {
            if (item?.folder && item.id) nextLevel.push(item.id);
          });
        });
      }
      level = nextLevel;
    }
    if (this._subtreePrefetchAbort === abortController) {
      this._subtreePrefetchAbort = null;
    }
  }

  /** Cuts short any in-flight _prefetchSiteSubtree walk. Call this whenever
   * Folder view is about to need a specific folder fetched right away — its
   * one request should never have to wait behind the background walk's
   * requests for the browser's per-origin connection slots. This is the fix
   * for "Folder view stays on Loading contents of X... for a long time while
   * List view comes back immediately": List view triggers the walk but never
   * awaits it (see openListViewFromFolderContext in DocumentsPage.tsx), so it
   * always renders straight away; the walk then keeps running in the
   * background and, without this cancellation, its own outstanding requests
   * were the thing actually blocking Folder view's next fetch. */
  public _cancelSiteSubtreePrefetch(): void {
    this._subtreePrefetchAbort?.abort();
    this._subtreePrefetchAbort = null;
  }

  /** Live SharePoint Term Store vessel names for a site, merged with DB
   * vessels server-side (see GET /api/sites/{site_id}/term-store-vessels).
   * Same lazy load-and-cache-then-forceUpdate pattern as
   * _getOrLoadSiteDrives, keyed per site so switching sites picks up that
   * site's own term store on next render. */
  public _getOrLoadTermStoreVessels(siteId: string): string[] {
    if (!siteId) return [];
    if (this._termStoreVesselsCache.has(siteId)) {
      return this._termStoreVesselsCache.get(siteId) || [];
    }
    this._termStoreVesselsCache.set(siteId, []);
    this._loadTermStoreVessels(siteId).then(names => {
      this._termStoreVesselsCache.set(siteId, names);
      this._scheduleForceUpdate();
    }).catch(() => undefined);
    return [];
  }

  public _getOrLoadSiteDrives(siteId: string): Array<{ id: string; name: string; web_url?: string; is_system?: boolean; item_count?: number | null }> | null {
    if (!siteId) return [];
    if (this._siteDrivesCache.has(siteId)) {
      return this._siteDrivesCache.get(siteId) || [];
    }
    this._siteDrivesCache.set(siteId, []);
    this._loadSiteDrives(siteId).then(drives => {
      this._siteDrivesCache.set(siteId, drives);
      this._scheduleForceUpdate();
    }).catch(() => undefined);
    return null;
  }

  public _getOrLoadSiteFolderChildren(
    siteId: string,
    driveId: string,
    folderId: string = 'root',
    force: boolean = false,
    signal?: AbortSignal,
  ): { items: any[]; loading?: boolean; error?: boolean; throttled?: boolean; retryAt?: number } {
    // Guard: cannot load without a valid site and drive
    if (!siteId || !driveId) {
      return { items: [], loading: false };
    }
    const key = `${siteId}::${driveId}::${folderId}`;
    if (force) {
      this._siteFolderItemsCache.delete(key);
    } else if (this._siteFolderItemsCache.has(key)) {
      const hit = this._siteFolderItemsCache.get(key)!;
      // A listing that failed only because Graph was throttling is not a real
      // result — once the cooldown ends, load it again instead of showing
      // "Couldn't load this folder" until the user presses Retry.
      if (!(hit.error && hit.retryAt && Date.now() >= hit.retryAt)) return hit;
      this._siteFolderItemsCache.delete(key);
    }

    // The recursive live-tree load already contains every folder and file in
    // the active drive. Use it for drill-down immediately instead of issuing
    // another Graph request with a potentially stale cached folder ID.
    if (!force && this.state.activeDocumentSite && this.state.documentLiveFolders.length > 0) {
      const liveItems = folderId === 'root'
        ? this.state.documentLiveFolders.filter(item => item.is_folder !== false && item.depth === 0)
        : this.state.documentLiveFolders.filter(item => item.parent_id === folderId);
      if (liveItems.length > 0 || folderId === 'root') {
        const items = liveItems.map(item => ({
          id: item.id,
          name: item.name,
          folder: item.is_folder !== false ? (() => {
            const children = this.state.documentLiveFolders.filter(child => child.parent_id === item.id);
            const prefix = `${item.path}/`;
            const descendants = this.state.documentLiveFolders.filter(child => child.path.startsWith(prefix));
            return {
              childCount: descendants.length,
              directFileCount: children.filter(child => child.is_folder === false).length,
              directFolderCount: children.filter(child => child.is_folder !== false).length,
            };
          })() : undefined,
          folder_counts: item.is_folder !== false ? (() => {
            const children = this.state.documentLiveFolders.filter(child => child.parent_id === item.id);
            const directFiles = children.filter(child => child.is_folder === false).length;
            const directFolders = children.filter(child => child.is_folder !== false).length;
            const prefix = `${item.path}/`;
            const totalFiles = this.state.documentLiveFolders.filter(child => child.path.startsWith(prefix) && child.is_folder === false).length;
            const totalFolders = this.state.documentLiveFolders.filter(child => child.path.startsWith(prefix) && child.is_folder !== false).length;
            return {
              direct_files: directFiles,
              direct_subfolders: directFolders,
              total_files: totalFiles,
              total_subfolders: totalFolders,
            };
          })() : undefined,
          file: item.is_folder === false ? {} : undefined,
          size: item.size,
          lastModifiedDateTime: item.last_modified_date_time,
          webUrl: item.web_url,
        }));
        const immediate = { items, loading: false, parentPath: '' };
        this._siteFolderItemsCache.set(key, immediate);
        return immediate;
      }
    }

    // Folder view is actively waiting on this one folder — free up the
    // browser's per-origin connection slots from any background List-view
    // subtree walk so this request isn't queued behind it (see
    // _cancelSiteSubtreePrefetch for the full story).
    this._cancelSiteSubtreePrefetch();

    const entry: { items: any[]; loading: boolean; _promise?: Promise<any[]> } = { items: [], loading: true };
    this._siteFolderItemsCache.set(key, entry);
    const loadPromise = this._loadSiteFolderChildren(siteId, driveId, folderId, signal).then(res => {
      if (signal?.aborted) {
        // Superseded (e.g. the user switched to a different site before this
        // resolved) — drop the placeholder instead of caching it as either
        // loaded or errored, same as _loadAndCacheSiteFolderChildren does,
        // so a later real request for this folder retries cleanly rather
        // than showing a stuck error/empty state.
        this._siteFolderItemsCache.delete(key);
        return;
      }
      const loadedEntry: { items: any[]; loading: boolean; parentPath: string; error: boolean; throttled?: boolean; retryAt?: number; errorAt?: number } = { items: res.items || [], loading: false, parentPath: res.parentPath || '', error: res.error === true };
      if (res.throttled) {
        loadedEntry.throttled = true;
        loadedEntry.retryAt = Math.max(this._graphThrottledUntil, Date.now() + 1000);
        window.setTimeout(() => this._scheduleForceUpdate(), Math.max(0, loadedEntry.retryAt - Date.now()) + 250);
      } else if (loadedEntry.error) {
        // Timed out / network blip / 5xx: retry by itself a few times.
        const retryIn = this._folderListingRetryDelayMs(key, true);
        if (retryIn > 0) {
          loadedEntry.retryAt = Date.now() + retryIn;
          window.setTimeout(() => this._scheduleForceUpdate(), retryIn + 250);
        }
      } else {
        this._folderListingRetryDelayMs(key, false);
      }
      if (loadedEntry.error) loadedEntry.errorAt = Date.now();
      this._siteFolderItemsCache.set(key, loadedEntry);
      if (res.folderId && res.folderId !== folderId) {
        const resolvedKey = `${siteId}::${driveId}::${res.folderId}`;
        this._siteFolderItemsCache.set(resolvedKey, loadedEntry);
      }
      // Auto-drill through a single-sub-folder, no-files listing (see
      // _autoDrilledFolderIds above) — but only when this is the folder the
      // user is actually looking at right now (the leaf of folderPathStack),
      // not a background/prefetch load for some other folder, and only once
      // per folder id so a deliberate Back navigation into it isn't fought.
      const currentStack = this.state.folderPathStack || [];
      const currentLeaf = currentStack[currentStack.length - 1];
      if (
        !loadedEntry.error &&
        currentLeaf && currentLeaf.id === folderId &&
        !this._autoDrilledFolderIds.has(folderId)
      ) {
        const childFolders = loadedEntry.items.filter((it: any) => !!it.folder);
        const childFiles = loadedEntry.items.filter((it: any) => !it.folder);
        if (childFolders.length === 1 && childFiles.length === 0) {
          this._autoDrilledFolderIds.add(folderId);
          const only = childFolders[0];
          this._pushFolderNav([...currentStack, { id: only.id, name: only.name }], this.state.docMainFolder);
          return;
        }
      }
      this._scheduleForceUpdate();
    }).catch(() => {
      if (signal?.aborted) {
        this._siteFolderItemsCache.delete(key);
      } else {
        this._siteFolderItemsCache.set(key, { items: [], loading: false, parentPath: '', error: true, errorAt: Date.now() });
      }
      this._scheduleForceUpdate();
    });
    // Lets _loadAndCacheSiteFolderChildren (used by the async path resolvers)
    // join this request instead of starting a second one and overwriting the
    // entry — the second one failing used to replace a good listing with an
    // error entry.
    entry._promise = loadPromise.then(() => ((this._siteFolderItemsCache.get(key)?.items) || []));
    return entry;
  }

  public async _refreshSiteFolder(siteId: string, driveId: string, folderId: string = 'root'): Promise<void> {
    if (!siteId || !driveId) return;
    const key = `${siteId}::${driveId}::${folderId}`;

    // This fires on every breadcrumb click, including Back/Forward and
    // re-clicking an earlier breadcrumb segment to step back to a folder
    // already open this session. It used to delete the cache entry first
    // (via _getOrLoadSiteFolderChildren(..., true)), which blanked the
    // already-loaded folder contents and put the "Loading contents of
    // X..." spinner back up — a real, noticeable reload — even though
    // nothing about that folder had changed. Keep whatever's already
    // cached on screen and refresh it in the background instead
    // (stale-while-revalidate): _getOrLoadSiteFolderChildren's render path
    // only shows the spinner when items.length === 0, so as long as the
    // previous items stay in the cache entry while loading is refreshed,
    // stepping back shows the folder instantly and swaps in fresh data
    // silently once it arrives.
    const existing = this._siteFolderItemsCache.get(key);
    if (existing && !existing.loading) {
      this._siteFolderItemsCache.set(key, { ...existing, loading: true });
    }

    // Same reasoning as _getOrLoadSiteFolderChildren: don't let a background
    // List-view subtree walk hold onto connection slots this foreground
    // refresh needs.
    this._cancelSiteSubtreePrefetch();

    this._loadSiteFolderChildren(siteId, driveId, folderId).then(res => {
      if (res.error) {
        // Keep a good listing on screen; for a throttled failure re-load by
        // itself once the cooldown ends (Retry used to leave an error entry
        // with no retryAt, so the folder stayed on "Couldn't load" for good).
        const retryAt = res.throttled ? Math.max(this._graphThrottledUntil, Date.now() + 1000) : undefined;
        const base = existing && !existing.error ? existing : { items: [], parentPath: '', error: true };
        this._siteFolderItemsCache.set(key, { ...base, loading: false, throttled: !!res.throttled, retryAt: base.error ? retryAt : undefined, errorAt: Date.now() });
        if (retryAt) window.setTimeout(() => this._scheduleForceUpdate(), retryAt - Date.now() + 250);
        this._scheduleForceUpdate();
        return;
      }
      const loadedEntry = { items: res.items || [], loading: false, parentPath: res.parentPath || '', error: false };
      this._siteFolderItemsCache.set(key, loadedEntry);
      this._scheduleForceUpdate();
    }).catch(() => {
      // A failed background refresh shouldn't wipe out contents that were
      // already showing correctly — keep them rather than falling back to
      // an error/empty state.
      if (existing) {
        this._siteFolderItemsCache.set(key, { ...existing, loading: false });
      } else {
        this._siteFolderItemsCache.set(key, { items: [], loading: false, parentPath: '', error: true });
      }
      this._scheduleForceUpdate();
    });
    this._scheduleForceUpdate();
  }

  private _usesBackendDocumentSite(): boolean {
    return false;
  }

  /** site_key of the site a "SharePoint Sites > <site> > <library> > …"
   * breadcrumb points at, or null for any other breadcrumb shape. */
  public _siteKeyForFolderStack(
    stack: { id: string; name: string }[],
    sites: Array<{ site_key: string; sp_site_name: string; site_id: string }> = this.state.documentSites,
  ): string | null {
    if (!stack || stack.length < 2 || stack[0]?.id !== 'sites_root') return null;
    const siteNode = stack[1];
    const rawId = String(siteNode?.id || '').replace(/^site:/, '');
    const name = String(siteNode?.name || '').trim().toLowerCase();
    const match = (sites || []).find(site =>
      (!!rawId && (site.site_id === rawId || site.site_key === rawId)) ||
      (!!name && ((site.sp_site_name || '').trim().toLowerCase() === name || (site.site_key || '').trim().toLowerCase() === name))
    );
    return match ? match.site_key : null;
  }

  /** Point this session's server-side active site at `siteKey` WITHOUT
   * touching the folder view (unlike _switchDocumentSite, which resets the
   * breadcrumb to the site's library root). */
  public async _syncSessionActiveSite(siteKey: string): Promise<void> {
    const base = this._base();
    const response = await fetch(`${base}/api/sites/active`, {
      method: 'POST',
      headers: {
        ...(base.includes('localhost') ? this._headersForLocalFallback() : this._headers()),
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ site_name: siteKey }),
    });
    if (!response.ok) throw new Error(`Could not set active site (${response.status})`);
  }

  public async _switchDocumentSite(siteKey: string): Promise<void> {
    const base = this._base();
    const selectedSite = this.state.documentSites.find(site => site.site_key === siteKey);
    if (!selectedSite) throw new Error(`SharePoint site '${siteKey}' is not available.`);

    // Free up the browser's per-origin connection slots from any background
    // subtree walk left over from the site being switched away from — its
    // requests are for a site we're about to stop showing, and would
    // otherwise compete with this switch's own requests for a connection
    // (same reasoning as _getOrLoadSiteFolderChildren / _refreshSiteFolder).
    this._cancelSiteSubtreePrefetch();

    // Cancel any still-running earlier call to THIS function — none of the
    // three requests below (the active-site POST, the drives lookup, the
    // root `/children` fetch at the end) used to be cancellable, so picking
    // a few sites from the dropdown in quick succession left every earlier
    // pick's requests running to completion in the background, all
    // competing with the latest pick's own requests for the browser's small
    // per-origin connection pool. That's what showed up as a pile-up of
    // "children" calls in the Network tab and a "Loading contents of
    // Documents..." that took a long time (or never finished) after
    // clicking through a few sites in a row. `generation` lets this call
    // recognize it's been superseded and bail out quietly — including
    // skipping the `throw` a plain abort would otherwise cause, which would
    // surface as a spurious "Could not switch site" alert (DocumentsPage.tsx
    // wraps every call to this function in `.catch(error => alert(...))`).
    this._siteSwitchAbort?.abort();
    const switchAbort = new AbortController();
    this._siteSwitchAbort = switchAbort;
    const generation = ++this._siteSwitchGeneration;
    const superseded = (): boolean => switchAbort.signal.aborted || generation !== this._siteSwitchGeneration;

    // The POST below only tells the backend which site is "active" for this
    // session; it doesn't return anything this call needs to keep going, and
    // resolving the site's live drives (below) only needs selectedSite.site_id,
    // which is already known. These two requests were previously awaited one
    // after the other — a second full network round-trip queued behind the
    // first for no reason other than being written sequentially. Firing them
    // together roughly halves the network wait before this site's folders
    // can start loading.
    const activeSitePromise = fetch(`${this._base()}/api/sites/active`, {
      method: 'POST',
      headers: {
        ...(base.includes('localhost') ? this._headersForLocalFallback() : this._headers()),
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ site_name: siteKey }),
      signal: switchAbort.signal,
    });

    // Site drives rarely change — once resolved for a site this session,
    // reuse that instead of re-fetching every time the user switches back to
    // a site they've already visited.
    const cachedDrives = this._siteDrivesCache.get(selectedSite.site_id);
    const drivesPromise = cachedDrives && cachedDrives.length > 0
      ? Promise.resolve(cachedDrives)
      : this._loadSiteDrives(selectedSite.site_id, switchAbort.signal);

    let response: Response;
    let liveDrives: Array<{ id: string; name: string; web_url?: string; is_system?: boolean; item_count?: number | null }>;
    try {
      [response, liveDrives] = await Promise.all([activeSitePromise, drivesPromise]);
    } catch (err: any) {
      if (superseded() || err?.name === 'AbortError') return;
      throw err;
    }
    if (superseded()) return;
    if (!response.ok) throw new Error(`Could not switch site (${response.status})`);
    // _siteFolderItemsCache entries are keyed `${siteId}::${driveId}::${folderId}`
    // (see _getOrLoadSiteFolderChildren), so they're already scoped to one
    // site's data — clearing the WHOLE cache here on every switch used to
    // force a full re-fetch of a site's folders even when switching straight
    // back to a site already browsed this session. That's the other half of
    // what made repeatedly toggling between a couple of sites pile up so
    // many "children" calls: every single toggle re-fetched everything from
    // scratch instead of the switch-back being instant. The `force: true`
    // passed to _getOrLoadSiteFolderChildren below already guarantees a
    // fresh look at the entering site's root; nothing here needs to touch
    // any other site's cached entries.
    let resolvedDriveId = selectedSite.drive_id;
    let resolvedLibraryName = selectedSite.default_library_name || 'Documents';
    try {
      if (liveDrives.length > 0) {
        this._siteDrivesCache.set(selectedSite.site_id, liveDrives);
        const preferredDrive = liveDrives.find(d => /^(documents|shared documents)$/i.test(d.name.trim()));
        const configuredDrive = liveDrives.find(d => d.id === selectedSite.drive_id);
        // Backend now flags SharePoint's own auto-provisioned libraries
        // (Site Assets, Style Library, Form Templates, ...) as `is_system`.
        // Falling back to plain `liveDrives[0]` used to mean a site with no
        // library literally named "Documents"/"Shared Documents" and no
        // `drive_id` configured yet could silently land the whole Documents
        // module on one of those system libraries — always empty, with
        // nothing in the UI to say why. Prefer the first non-system drive
        // (ideally one that actually has items) before falling back to
        // whatever Graph returned first.
        const nonSystemDrives = liveDrives.filter(d => !d.is_system);
        const bestNonSystemDrive = nonSystemDrives.find(d => (d.item_count ?? 0) > 0) || nonSystemDrives[0];
        const activeDrive = preferredDrive || configuredDrive || bestNonSystemDrive || liveDrives[0];
        resolvedDriveId = activeDrive.id;
        resolvedLibraryName = activeDrive.name || resolvedLibraryName;
      }
    } catch (driveError) {
      console.warn('[VesselDMS] Could not resolve live document drive while switching sites:', driveError);
    }
    const siteName = selectedSite.sp_site_name || selectedSite.site_key;
    const libraryNode = { id: `drive:${resolvedDriveId}`, name: resolvedLibraryName };
    const siteNode = { id: `site:${selectedSite.site_id}`, name: siteName };
    const selectedRoot = [
      { id: 'sites_root', name: 'SharePoint Sites' },
      siteNode,
      libraryNode,
    ];
    // Previously this also called `await this._loadData(true)` here, then
    // re-applied this exact setState a second time afterwards to "restore"
    // the breadcrumb root because that reload's async vessel/rows updates
    // could race and clobber it. Two problems with that:
    //  1. `_loadData` re-fetches `/api/vessels` with no site_key (see
    //     backend/app/services/real_backend.py:838-859 `list_vessels`),
    //     which returns vessels from EVERY configured site, not just the one
    //     just switched to — populating `this.state.vessels` with
    //     other-site vessels while the user is looking at this site's
    //     folders. `_loadData`'s own fallback walk (`_flattenAll`) can also
    //     fire the same "children" call fan-out described in
    //     `_refreshSingleVesselRows`'s comment above.
    //  2. It is entirely wasted work: `docScopeType: 'sites'` means the
    //     vessel-based Documents table isn't even the active view.
    // Dropping the full reload removes both the redundant call burst and the
    // window where other-site vessel data could leak into state while
    // browsing this site's folders. The direct, explicitly-scoped
    // `_getOrLoadSiteFolderChildren(selectedSite.site_id, selectedSite.drive_id, ...)`
    // call below is the only fetch this view actually needs.
    this.setState({
      activeDocumentSite: siteKey,
      documentSites: this.state.documentSites.map(site => site.site_key === siteKey
        ? { ...site, drive_id: resolvedDriveId, default_library_name: resolvedLibraryName }
        : site),
      documentLiveFolders: [],
      documentLiveFoldersLoading: false,
      rows: [],
      uploadedFilesByFolder: {},
      docScopeType: 'sites',
      docViewMode: 'folder',
      folderPathStack: selectedRoot,
      docMainFolder: 'SharePoint Sites',
      folderNavHistory: [{ folderPathStack: selectedRoot, docMainFolder: 'SharePoint Sites' }],
      folderNavIndex: 0,
      // Vessel options are site-specific; a vessel picked on the previous
      // site must not stay selected after switching.
      vesselFilter: 'all',
      // Main folder / sub-folder / category filters point at folder NAMES
      // from whichever site was active before this switch. Those names
      // don't necessarily exist in the new site's folder tree — leaving
      // them set stranded the Main folder dropdown on a name the new site
      // has no match for, which made the Sub-folder dropdown compute zero
      // options and show the disabled "No folder found" state even though
      // the new site does have sub-folders, and could also drive the
      // stuck-drive-root folder id resolution off the wrong path. Reset
      // every one of these alongside vesselFilter so the new site starts
      // from "All main folders" / "All sub-folders", same as opening
      // Documents fresh.
      docGroupFilter: 'all',
      docCategoryFilter: 'all',
      docGroupLevelFilter: 'all',
      docLeafCategoryFilter: 'all',
      docSubCategoryFilter: 'all',
      docSubfolderOtherFilter: 'all',
      catFilter: 'all',
      textFilter: '',
      docListPage: 0,
    });
    this._getOrLoadSiteFolderChildren(selectedSite.site_id, resolvedDriveId, 'root', true, switchAbort.signal);
    // Kick off the Home/search dashboard's per-site scan (real_backend.py
    // _dashboard_site_scan) in the background as soon as the switch lands,
    // instead of waiting for the user's first keystroke in the Documents
    // search box to trigger it (see _triggerGlobalSearch's call to
    // /api/dashboard/documents). That scan can take a while the first time
    // a site is opened this session — especially when the site's SharePoint
    // search index lags and the backend falls back to a recursive folder
    // walk — and starting it now overlaps that wait with the time the user
    // spends looking at the folder view that just loaded. A search fired a
    // few seconds later then hits an already-warm (or already in-flight)
    // cache instead of a cold scan_pending=true start, which is what made
    // "switch site, then search" feel much slower than searching again on
    // a site already browsed this session. page_size=5 (the backend's
    // minimum — see page_size: Query(..., ge=5) in dashboard_documents)
    // keeps this request as cheap as the API allows — its only job is to
    // populate the shared per-site cache that _triggerGlobalSearch and the
    // Home dashboard both read from. page_size=1 used to be passed here,
    // but that fails backend validation (422) and made this priming call a
    // silent no-op since the response was never awaited/read.
    this._fetchJson(
      `${this._base()}/api/dashboard/documents?site_key=${encodeURIComponent(siteKey)}&page=1&page_size=5`,
      switchAbort.signal
    ).catch(() => undefined);
    if (this._siteSwitchAbort === switchAbort) {
      this._siteSwitchAbort = null;
    }
  }

  // ── Alert Bell (top-header) ────────────────────────────────────────────────
  // Fetches the new folder-creation alerts (newly created SharePoint Online
  // folders + newly provisioned vessels) and surfaces them under the header bell.

  public _fetchAlerts = (signal?: AbortSignal): void => {
    this._fetchJson(`${this._base()}/api/alerts/all`, signal)
      .then((res: any) => {
        if (Array.isArray(res)) {
          // Merge backend alerts with locally-generated ones (vessel_unrecognised, vessel_deleted)
          // so that SPO delta-sync alerts are not wiped out by the periodic backend poll.
          this.setState(prev => {
            const backendIds = new Set((res as AlertItem[]).map(a => a.id));
            const localOnly = prev.alertsList.filter(
              a => !backendIds.has(a.id) &&
                (a.alert_type === 'vessel_unrecognised' || a.alert_type === 'file_outside_structure' || a.alert_type === 'vessel_deleted' || a.alert_type === 'document_deleted')
            );
            // Once an alert is marked read locally (optimistic update in
            // _markAlertRead / _markAllAlertsRead), keep it read even if this
            // poll (every 30s) lands before the backend's own POST .../read
            // has been persisted. Without this, the top-bar bell badge could
            // briefly revert to unread/reappear after the person had already
            // dismissed it.
            const prevReadById = new Map(prev.alertsList.map(a => [a.id, a.read]));
            const merged = (res as AlertItem[]).map(a =>
              prevReadById.get(a.id) === true && !a.read ? { ...a, read: true } : a,
            );
            return { alertsList: [...localOnly, ...merged], alertsLoaded: true };
          });
        } else {
          this.setState({ alertsLoaded: true });
        }
      })
      .catch(() => undefined);
  };

  /** Who/where a deleted item came from — resolved from the backend's
   *  deletion_log (via GET /api/recycle-bin/nodes), not guessed on the client. */
  public _lookupDeletionAttribution = async (
    nodeId: string,
    nameHint?: string,
  ): Promise<{ deletedByEmail: string | null; deletedByName: string | null; siteName: string | null; originalPath: string | null } | null> => {
    try {
      const items = await this._fetchJson(`${this._base()}/api/recycle-bin/nodes`);
      const list = Array.isArray(items) ? items : [];
      const hit =
        list.find((x: any) => String(x?.id || '') === String(nodeId)) ||
        (nameHint ? list.find((x: any) => String(x?.name || '').toLowerCase() === nameHint.toLowerCase()) : undefined);
      if (!hit) return null;
      return {
        deletedByEmail: hit.deleted_by_email || null,
        deletedByName: hit.deleted_by_name || null,
        siteName: hit.site_name || null,
        originalPath: hit.original_path || hit.spo_path || null,
      };
    } catch {
      return null;
    }
  };

  /** Add a local activity alert and immediate recycle-bin popup for a deleted SPO item.
   *  `attribution`, when already known (e.g. from the recycle-bin feed), is used
   *  immediately; otherwise this fires an async lookup and patches the alert +
   *  toast in place once the backend resolves who deleted it, from where, and
   *  on which site — so the popup never falsely credits the viewing user. */
  public _handleSpoDocumentDeletion(
    node: import('./deltaSync').SpoFolderNode,
    attribution?: { deletedByEmail?: string | null; deletedByName?: string | null; siteName?: string | null; originalPath?: string | null },
  ): void {
    // Already raised a popup for this exact SPO item — most commonly the
    // root folder of a vessel that was already reported via
    // _handleSpoVesselDeletion, or a delete already shown once from a
    // retried scan. Skip rather than raise a second, duplicate popup.
    if (this._shownDeletionToastIds.has(node.id)) return;
    this._shownDeletionToastIds.add(node.id);

    const now = new Date().toISOString();
    const itemType = node.isFolder ? 'folder' : 'file';
    const activeSiteObj = (this.state.documentSites || []).find(s => s.site_key === this.state.activeDocumentSite);
    const path = attribution?.originalPath || node.serverRelativePath || 'SharePoint Online Documents';
    const siteName = attribution?.siteName || activeSiteObj?.sp_site_name || this.props.siteUrl || null;
    const deletedByName = attribution?.deletedByName || null;
    const deletedByEmail = attribution?.deletedByEmail || null;

    const alert: AlertItem = {
      id: `document_deleted_${node.id}_${Date.now()}`,
      drive_item_id: node.id,
      folder_name: node.name,
      folder_path: path,
      parent_folder_id: node.parentId,
      vessel_name: null,
      department: 'Documents',
      created_by_email: deletedByEmail || '',
      created_by_name: deletedByName || 'Unknown user',
      alert_type: 'document_deleted',
      alert_category: 'dms',
      read: false,
      created_at: now,
      item_type: itemType,
      spo_path: path,
      site_name: siteName,
    };

    let inserted = false;
    this.setState(prev => {
      const alreadyTracked = prev.alertsList.some(
        a => a.alert_type === 'document_deleted' && a.drive_item_id === node.id,
      );
      if (alreadyTracked) return null;
      inserted = true;
      return {
        alertsList: [alert, ...prev.alertsList],
        spoDocumentDeletedToast: {
          itemType,
          items: [
            ...(prev.spoDocumentDeletedToast?.items || []),
            { id: node.id, name: node.name, path, siteName, deletedByName, deletedByEmail },
          ],
        },
      };
    });
    if (inserted) this._scheduleDocumentToastAutoClose();

    // If we don't yet know who deleted this (the common case: the frontend
    // only sees the delta-sync "it's gone" event, not the actor), ask the
    // backend's recycle-bin feed — it's enriched from SharePoint's own
    // recycle bin (DeletedByEmail/DeletedByName) once that's propagated.
    if (inserted && !deletedByEmail && !deletedByName) {
      void this._lookupDeletionAttribution(node.id, node.name).then(resolved => {
        if (!resolved) return;
        this.setState(prev => ({
          alertsList: prev.alertsList.map(a =>
            a.alert_type === 'document_deleted' && a.drive_item_id === node.id
              ? {
                  ...a,
                  created_by_email: resolved.deletedByEmail || a.created_by_email,
                  created_by_name: resolved.deletedByName || a.created_by_name,
                  site_name: resolved.siteName || a.site_name,
                  folder_path: resolved.originalPath || a.folder_path,
                }
              : a,
          ),
          spoDocumentDeletedToast: prev.spoDocumentDeletedToast && {
            ...prev.spoDocumentDeletedToast,
            items: prev.spoDocumentDeletedToast.items.map(entry =>
              entry.id === node.id
                ? {
                    ...entry,
                    path: resolved.originalPath || entry.path,
                    siteName: resolved.siteName || entry.siteName,
                    deletedByName: resolved.deletedByName || entry.deletedByName,
                    deletedByEmail: resolved.deletedByEmail || entry.deletedByEmail,
                  }
                : entry,
            ),
          },
        }));
        if (this.state.spoDocumentDeletedToast) this._scheduleDocumentToastAutoClose();
      });
    }
  }

  /**
   * Fallback for deleted items not present in local spoFolderMap.
   * Resolves details from backend recycle-bin feed and still emits popup alert.
   */
  public _handleUnknownSpoDeletion = (deletedId: string): void => {
    this._fetchJson(`${this._base()}/api/recycle-bin/nodes`)
      .then((items: any) => {
        const list = Array.isArray(items) ? items : [];
        const hit = list.find((x: any) => String(x?.id || '') === String(deletedId));

        if (hit) {
          const kind = String(hit.kind || '').toLowerCase();
          const isFolder = kind !== 'file';
          const node = {
            id: deletedId,
            name: (hit.name || 'SharePoint item').toString(),
            parentId: null,
            isFolder,
            serverRelativePath: (hit.original_path || hit.spo_path || 'SharePoint Online Documents').toString(),
            children: [],
          } as import('./deltaSync').SpoFolderNode;
          this._handleSpoDocumentDeletion(node, {
            deletedByEmail: hit.deleted_by_email || null,
            deletedByName: hit.deleted_by_name || null,
            siteName: hit.site_name || null,
            originalPath: hit.original_path || hit.spo_path || null,
          });
          return;
        }

        const node = {
          id: deletedId,
          name: 'SharePoint folder',
          parentId: null,
          isFolder: true,
          serverRelativePath: 'SharePoint Online Documents',
          children: [],
        } as import('./deltaSync').SpoFolderNode;
        this._handleSpoDocumentDeletion(node);
      })
      .catch(() => {
        const node = {
          id: deletedId,
          name: 'SharePoint folder',
          parentId: null,
          isFolder: true,
          serverRelativePath: 'SharePoint Online Documents',
          children: [],
        } as import('./deltaSync').SpoFolderNode;
        this._handleSpoDocumentDeletion(node);
      });
  };

  public _unreadAlertCount(): number {
    return this.state.alertsList.filter(a => !a.read).length;
  }

  // The bell no longer opens a popup: it opens the Notifications page.
  public _toggleAlertBell = (): void => {
    this._fetchAlerts();
    this._openAlertsPage();
  };

  public _closeAlertBell = (): void => {
    this.setState({ alertOpen: false });
  };

  public _markAlertRead = (id: string): void => {
    this.setState(prev => ({
      alertsList: prev.alertsList.map(a => (a.id === id ? { ...a, read: true } : a)),
    }));
    if (id.startsWith('anomaly_')) {
      const anomalyId = id.slice('anomaly_'.length);
      if (/^\d+$/.test(anomalyId)) {
        this._fetchJson(`${this._base()}/api/anomalies/${anomalyId}/read`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ read: true }),
        }).catch(() => undefined);
      }
      return;
    }
    // Local activity/email alerts do not have FolderAlert database IDs.
    if (!/^\d+$/.test(id)) return;
    this._fetchJson(`${this._base()}/api/alerts/${id}/read`, { method: 'POST' }).catch(() => undefined);
  };

  public _markAllAlertsRead = (): void => {
    this.setState(prev => ({ alertsList: prev.alertsList.map(a => ({ ...a, read: true })) }));
    this._fetchJson(`${this._base()}/api/alerts/read-all`, { method: 'POST' }).catch(() => undefined);
  };

  public _setAlertFilter = (filter: 'all' | 'unread'): void => {
    this.setState({ alertFilter: filter });
  };

  public _setAlertCategory = (category: 'all' | 'dms' | 'crud' | 'email'): void => {
    this.setState({ alertCategory: category });
  };

  public _openAlertsPage = (alertId: string | null = null): void => {
    if (alertId) this._markAlertRead(alertId);
    this.setState({ selectedAlertId: alertId, alertOpen: false, view: 'alerts' });
  };

  /**
   * Convert current folderAnomalies into AlertItems and merge them into
   * alertsList so they appear in the header bell without duplicates.
   * Called after every anomaly scan / setState that updates folderAnomalies.
   */
  public _syncAnomalyAlerts(): void {
    const { folderAnomalies, alertsList } = this.state;
    if (!folderAnomalies || folderAnomalies.length === 0) return;

    const existingAnomalyIds = new Set(
      alertsList
        .filter(a => a.anomaly_id !== undefined)
        .map(a => a.anomaly_id)
    );

    const newAlerts: AlertItem[] = folderAnomalies
      .filter(a => !existingAnomalyIds.has(a.id))
      .map(a => {
        let alertType: AlertItem['alert_type'];
        if (a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'folder') {
          alertType = 'vessel_unrecognised';
        } else if (a.item_type === 'file') {
          alertType = 'file_outside_structure';
        } else {
          alertType = 'subfolder_anomaly';
        }
        return {
          id: `anomaly_${a.id}`,
          drive_item_id: a.drive_item_id,
          folder_name: a.name,
          folder_path: a.spo_path,
          parent_folder_id: null,
          vessel_name: a.vessel_name,
          department: a.department,
          created_by_email: '',
          created_by_name: 'SharePoint Online',
          alert_type: alertType,
          read: a.read === true,
          read_at: a.read_at,
          created_at: a.detected_at,
          updated_at: a.updated_at,
          anomaly_id: a.id,
          item_type: a.item_type,
          spo_path: a.spo_path,
        } as AlertItem;
      });

    if (newAlerts.length > 0) {
      this.setState(prev => ({ alertsList: [...newAlerts, ...prev.alertsList] }));
    }
  }


  // ── Data Loading ──────────────────────────────────────────────────────────

  public async _loadData(force = false): Promise<void> {
    if (this._isLoadingData && !force) {
      return;
    }
    // Cancel any pending debounced reload since we're loading now
    if (this._deltaReloadTimer) {
      clearTimeout(this._deltaReloadTimer);
      this._deltaReloadTimer = null;
    }
    this._isLoadingData = true;
    this._fileFolderRefreshRequestedAt.clear();
    this._abort?.abort();
    this._abort = new AbortController();
    const signal = this._abort.signal;
    try {
      const base = this._base();
      if (!base) {
        const merged = this._mergeWithUploadCache([]);
        this.setState({ loading: false, rows: merged.rows, uploadedFilesByFolder: merged.uploadedFilesByFolder });
        return;
      }
      if (this.state.rows.length === 0) {
        this.setState({ loading: true, error: null });
      }

      // Kick off non-critical background tasks
      this._fetchAnomalies(signal);
      this._loadNormalFolders();
      void this._loadRecycleBin();
      void this._loadDashboardStats(); // eagerly fetch dashboard counts

      // Step 1: Always fetch vessel list from backend database first
      let vesselList: any = null;
      let vesselListFetched = false;
      // Hoisted out of the try block below (was previously try-scoped-only)
      // so Step 1a can also read it after the fetch.
      const siteFilterValue = this.state.vesselSiteFilter && this.state.vesselSiteFilter !== 'all'
        ? this.state.vesselSiteFilter
        : '';
      try {
        const vesselUrl = siteFilterValue
          ? `${base}/api/vessels?site_key=${encodeURIComponent(siteFilterValue)}`
          : `${base}/api/vessels`;
        vesselList = await this._fetchJson(vesselUrl, signal);
        vesselListFetched = true;
      } catch {
        /* ignore */
      }
      if (signal.aborted) return;

      // Step 1a: keep the vessel list current without a manual "Sync Vessels
      // from SharePoint" click — auto-reconcile root folders / Term Store
      // vessel terms in the background (throttled per site, see
      // _maybeAutoSyncVesselsFromSharePoint), then silently refresh once if
      // it actually found something new.
      if (vesselListFetched) {
        this._maybeAutoSyncVesselsFromSharePoint(siteFilterValue || null);
      }

      let vessels: VesselRecord[] = [];
      const { spoDeletedVesselIds, recycleBin } = this.state;
      const deletedNames = new Set(
        recycleBin
          .filter(r => r.kind === 'vessel' || r.item_type === 'vessel')
          .map(r => r.name.toLowerCase())
      );
      if (vesselList && Array.isArray(vesselList) && vesselList.length > 0) {
        vessels = this._filterDeletedVessels(
          vesselList.map((v: any) => ({ ...v, name: cleanName(v.name), status: v.status || 'Active' })),
          recycleBin,
          spoDeletedVesselIds,
          siteFilterValue || null,
        );
      }
      // Preserve any locally created vessels currently in state that might not yet be returned by backend
      const fetchedNames = new Set(vessels.map(v => (v.name || '').trim().toLowerCase()));
      // Only carry over optimistic, not-yet-persisted records. Rows that came
      // from an earlier backend fetch (tagged _fetched_for_site) are
      // superseded by this fetch — carrying them over is how a previously
      // selected site's vessels leaked into the newly selected site.
      const pendingLocalVessels = (this.state.vessels || []).filter(v =>
        !(v as any)._fetched_for_site &&
        !fetchedNames.has((v.name || '').trim().toLowerCase()) &&
        !spoDeletedVesselIds.has(v.id) &&
        !deletedNames.has((v.name || '').trim().toLowerCase())
      );
      if (pendingLocalVessels.length > 0) {
        vessels = [...pendingLocalVessels, ...vessels];
      }

      // Step 1b: SPO anomaly scan runs fully in the BACKGROUND — does NOT block vessel/document rendering
      const { graphClient, siteId, driveId } = this.props;
      // Keep the Documents list usable when the configured development API is
      // unavailable (such as an untrusted TLS certificate). In that case,
      // discover vessels from Documents / Vessels directly through Graph.
      // SPO fallback is only for backend-unavailable cases. If backend returns
      // an empty list, respect it (e.g., after deleting the last vessel)
      // and do not auto-recreate vessels from folder names.
      if (!vesselListFetched && vessels.length === 0 && graphClient && siteId && driveId) {
        // New structure: vessels are under each MainFolder, not a shared root.
        // Use Technical & Crewing as the primary source for vessel names.
        const vesselNodes = await this._getGraphChildren(this.MAIN_FOLDER_NAMES[0], signal).catch(() => []);
        if (!signal.aborted) {
          vessels = vesselNodes
            .filter(node => node.isFolder && node.name && !/^pool-/i.test(node.name)
              && node.name.toLowerCase() !== 'common for all ships'
              && node.name.toLowerCase() !== 'common (not ship specific)')
            .map(node => ({
              id: node.id || node.name,
              name: cleanName(node.name),
              status: 'Active',
              is_provisioned: true,
            }));
        }
      }
      if (graphClient && siteId && driveId) {
        void (async () => {
          try {
            const spoAnomalies: FolderAnomalyItem[] = [];

            // 1. Scan main root level
            const mainRootNodes = await this._getGraphChildren('', signal).catch(() => []);
            if (signal.aborted) return;
            const knownMainDepts = new Set(['commercial & chartering', 'insurance', 'kaizen - knowledge bank', 'technical & crewing', 'knowledge bank']);

            for (const node of mainRootNodes) {
              if (!node.name) continue;
              const cleanN = cleanName(node.name).trim().toLowerCase();
              if (knownMainDepts.has(cleanN)) continue;
              spoAnomalies.push({
                id: Date.now() + Math.floor(Math.random() * 100000),
                drive_item_id: node.id || `spo_main_${node.name}`,
                name: node.name,
                item_type: node.isFolder ? 'folder' : 'file',
                anomaly_type: 'main_folder_unmatched',
                department: 'Main Root',
                vessel_name: null,
                spo_path: node.name,
                resolved: false,
                detected_at: new Date().toISOString(),
              });
            }

            // 2. Scan Technical & Crewing level — only if that folder
            // actually exists at this site's root (step 1 already fetched
            // the full root listing). Sites with a genuinely different
            // top-level structure don't have it, and probing for it
            // unconditionally on every load/site-switch was a guaranteed,
            // repeated 404 for no benefit.
            const technicalCrewingNode = mainRootNodes.find(node => node.isFolder && cleanName(node.name).trim().toLowerCase() === 'technical & crewing');
            const spoVesselNodes = technicalCrewingNode
              ? await this._getGraphChildren(technicalCrewingNode.name, signal).catch(() => [])
              : [];
            if (signal.aborted) return;
            const stripVP = (name: string): string =>
              (name || '').trim().toLowerCase().replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '');
            const dbVNames = new Set(vessels.map(v => (v.name || '').trim().toLowerCase()));
            const dbVStripped = new Set(vessels.map(v => stripVP(v.name)));
            const knownCommon = new Set([
              ...MAIN_FOLDERS.flatMap(main => main.perVesselTree.map(folder => folder.name.toLowerCase())),
              'common for all ships',
              'common (for all ships)', 'common agreements (not ship specific)', 'common (not ship specific)',
              'common agreements', 'common'
            ]);
            for (const node of spoVesselNodes) {
              if (!node.name) continue;
              const cleanN = cleanName(node.name).trim().toLowerCase();
              const strippedN = stripVP(node.name);
              if (dbVNames.has(cleanN) || dbVStripped.has(strippedN) ||
                cleanN.includes('common') || knownCommon.has(cleanN) || cleanN.startsWith('pool-')) continue;
              spoAnomalies.push({
                id: Date.now() + Math.floor(Math.random() * 100000),
                drive_item_id: node.id || `spo_vessel_${node.name}`,
                name: node.name,
                item_type: node.isFolder ? 'folder' : 'file',
                anomaly_type: 'vessel_level_unmatched',
                department: 'Technical & Crewing',
                vessel_name: null,
                spo_path: `Technical & Crewing/${node.name}`,
                resolved: false,
                detected_at: new Date().toISOString(),
              });
            }
            if (spoAnomalies.length > 0) {
              this.setState(prev => {
                const existing = new Set(prev.folderAnomalies.map(a => a.spo_path));
                const newOnly = spoAnomalies.filter(a => !existing.has(a.spo_path));
                if (newOnly.length === 0) return null as any;
                return { folderAnomalies: [...prev.folderAnomalies, ...newOnly] };
              }, () => this._syncAnomalyAlerts());
            }
          } catch (spoErr) {
            console.warn('[VesselDMS] SPO anomaly scanning warning:', spoErr);
          }
        })();
      }

      // Update vessel list in state immediately so UI shows vessels right away
      if (vessels.length > 0) {
        // Deduplicate by name before setting state
        const seenNames = new Set<string>();
        vessels = vessels.filter(v => {
          const key = (v.name || '').trim().toLowerCase();
          if (seenNames.has(key)) return false;
          seenNames.add(key);
          return true;
        });
        this.setState({ vessels });
      }

      // Step 2: Fast path — backend flat-tree endpoint (single DB query, ~50ms)
      // This is always preferred over the Graph folder-walk for list view data.
      // The landing list/folder views show only the four newest vessels. A
      // specific vessel is fetched on demand by _loadVesselRowsFromApi.
      const flatTreeParams = new URLSearchParams({ vessel_limit: '4', vessel_offset: '0' });
      if (force) flatTreeParams.set('force_refresh', 'true');
      const flatTreeUrl = `${base}/api/vessels/flat-tree?${flatTreeParams.toString()}`;
      const flatTree = await this._fetchJson(flatTreeUrl, signal).catch(() => null);
      if (signal.aborted) return;
      if (flatTree && Array.isArray(flatTree) && flatTree.length > 0) {
        let initialRows = this._normalize(flatTree);
        const initialVessels = vessels.slice(0, 4);
        const rowVesselNames = new Set(initialRows.map(row => cleanName(row.vesselName).trim().toLowerCase()));

        // Some deployed backend versions ignore vessel_limit and return rows
        // for only one vessel. Fill any missing recent vessels with targeted
        // requests so the Documents list always starts with four vessels.
        const missingVessels = initialVessels.filter(v => !rowVesselNames.has(cleanName(v.name).trim().toLowerCase()));
        if (missingVessels.length > 0) {
          const supplements = await Promise.all(missingVessels.map(async vessel => {
            const url = `${this._base()}/api/vessels/flat-tree?vessel_name=${encodeURIComponent(vessel.name)}`;
            const data = await this._fetchJson(url, signal).catch(() => null);
            return Array.isArray(data) ? this._normalize(data) : [];
          }));
          initialRows = initialRows.concat(...supplements);
        }

        const initialVesselNames = initialVessels.map(vessel => vessel.name);
        // Deduplicate vessels before final state update
        const seenVesselNames2 = new Set<string>();
        vessels = vessels.filter(v => {
          const key = (v.name || '').trim().toLowerCase();
          if (seenVesselNames2.has(key)) return false;
          seenVesselNames2.add(key);
          return true;
        });
        const merged = this._mergeWithUploadCache(initialRows);
        this.setState({
          vessels,
          rows: merged.rows,
          uploadedFilesByFolder: merged.uploadedFilesByFolder,
          loading: false,
          documentVesselCount: Math.min(4, vessels.length),
          documentVesselsLoadingMore: false,
          // Show all four recent vessels on the initial Documents load.
          vesselFilter: 'all',
          docGroupFilter: 'all',
          catFilter: 'all',
          docListPage: 0,
        }, () => {
          // The flat tree contains folder structure only. Walk the active
          // SharePoint tree so parent folders with child folders (for example
          // Invoices & Payments > Invoice) are traversed to their file leaves.
          // Skipped while browsing a live-library scope ('sites' /
          // 'shared_docs' / 'documents'): _loadData runs unconditionally
          // (mount, an explicit Refresh click, reloadKey bumps, ...)
          // regardless of which Documents view is currently open, but this
          // walk resolves each DB row's folder via the legacy default
          // department template (Technical & Crewing/Vessels/<vessel>,
          // etc.), which doesn't match a real site's actual folder names and
          // just floods Graph with 404s — the file listing in those scopes
          // already comes from _getOrLoadSiteFolderChildren / the live
          // folder tree, not from DB rows.
          const isLiveLibraryScope = this.state.docScopeType === 'sites' ||
            this.state.docScopeType === 'shared_docs' || this.state.docScopeType === 'documents';
          if (!isLiveLibraryScope) {
            this._refreshFilesFromBackendRows(0);
          }
        });
        return;
      }

      // Step 3: Fallback — Graph API folder-walk (only when flat-tree returns nothing)
      if (graphClient && siteId && driveId && vessels.length > 0) {
        try {
          const graphRows = await this._flattenAllViaGraph(vessels, signal);
          if (!signal.aborted) {
            const merged = this._mergeWithUploadCache(this._normalize(graphRows));
            this.setState({ vessels, rows: merged.rows, uploadedFilesByFolder: merged.uploadedFilesByFolder, loading: false });
            return;
          }
        } catch (graphErr) {
          console.warn('[VesselDMS] _loadData: Graph walk failed, falling back to REST API:', graphErr);
        }
      }

      if (signal.aborted) return;

      // Step 4: Final fallback — REST API per-vessel/per-folder walk
      if (vessels.length > 0) {
        const rows = await this._flattenAll(vessels, signal);
        if (!signal.aborted) {
          const merged = this._mergeWithUploadCache(this._normalize(rows));
          this.setState({ vessels, rows: merged.rows, uploadedFilesByFolder: merged.uploadedFilesByFolder, loading: false });
        }
      } else {
        const merged = this._mergeWithUploadCache([]);
        this.setState({ vessels, rows: merged.rows, uploadedFilesByFolder: merged.uploadedFilesByFolder, loading: false });
      }
    } catch (err: any) {
      if (!signal.aborted) this.setState({ loading: false, error: err?.message ?? 'Failed to load data.' });
    } finally {
      this._isLoadingData = false;
    }
  }

  public _normalize(raw: FlatRow[]): FlatRow[] {
    return raw.map(r => ({ ...r, vesselName: cleanName(r.vesselName), group: cleanName(r.group), category: cleanName(r.category) }));
  }

  /** Convert the UI breadcrumb to its location in this site's Documents library. */
  public _sharePointFolderPath(subFolderPath: string, fallback = ''): string {
    const parts = (subFolderPath || '').split('>').map(part => part.trim()).filter(Boolean);
    const canonicalParts = parts.filter((part, index) =>
      index === 0 || part.toLowerCase() !== parts[index - 1].toLowerCase()
    );
    if (canonicalParts.length === 0) return fallback.replace(/^\/+/, '');

    const isMainAt1 = canonicalParts.length >= 2 && this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === canonicalParts[1].toLowerCase());
    const isMainAt0 = this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === canonicalParts[0].toLowerCase());

    if (isMainAt1 && !isMainAt0) {
      // Breadcrumb is "MV 124 > Technical & Crewing > Category..."
      // SharePoint folder is "Technical & Crewing/MV 124/Category..."
      return [canonicalParts[1], canonicalParts[0], ...canonicalParts.slice(2)].join('/');
    }

    if (!isMainAt0 && !/^kaizen/i.test(canonicalParts[0]) && !/^common/i.test(canonicalParts[0])) {
      // Breadcrumb is "MV 124 > Category..." without Main Folder
      const activeMain = this.state.docMainFolder || 'Technical & Crewing';
      return [activeMain, ...canonicalParts].join('/');
    }

    return canonicalParts.join('/');
  }


  /**
   * Immediately fetch rows for a specific vessel from the backend
   * (uses the vessel_name-filtered flat-tree endpoint, hits DB only, ~50ms).
   * Merges the result into `rows` so the list view updates right away.
   */
  public async _loadVesselRowsFromApi(vesselName: string): Promise<void> {
    const base = this._base();
    if (!base || !vesselName || vesselName === 'all') return;
    const normV = vesselName.trim().toLowerCase();
    this.setState({ vesselLoadingName: vesselName, documentFilesLoading: true });
    try {
      const url = `${base}/api/vessels/flat-tree?vessel_name=${encodeURIComponent(vesselName)}`;
      const data = await this._fetchJson(url, new AbortController().signal).catch(() => null);
      if (!data || !Array.isArray(data) || data.length === 0) {
        this.setState({ vesselLoadingName: null, documentFilesLoading: false });
        return;
      }
      const incoming: FlatRow[] = this._normalize(data);
      this.setState(prev => {
        const fileRowKey = (row: FlatRow): string =>
          [
            (row.subFolderPath || '').trim().toLowerCase(),
            (row.fileName || '').trim().toLowerCase(),
          ].join('||');
        const incomingFileKeys = new Set(
          incoming
            .filter(r => Boolean(r.fileName))
            .map(fileRowKey)
        );
        const preservedExistingFiles = prev.rows.filter(r =>
          (r.vesselName || '').trim().toLowerCase() === normV &&
          Boolean(r.fileName) &&
          !incomingFileKeys.has(fileRowKey(r))
        );
        // Replace/merge: remove old rows for this vessel, then prepend fresh ones
        const kept = prev.rows.filter(r => (r.vesselName || '').trim().toLowerCase() !== normV);
        return {
          rows: [...incoming, ...preservedExistingFiles, ...kept],
          vesselLoadingName: null,
          documentFilesLoading: false,
        };
      }, () => {
        // Reset the run-once guard so the new vessel's rows are always refreshed.
        this._lastRefreshedRowsKey = '';
        this._refreshFilesFromBackendRows();
        // DB rows describe the folder structure, but existing files live in
        // SharePoint. Walk the active drive so a refresh never depends on
        // stale cached folder IDs.
      });
    } catch (err) {
      console.warn('[VesselDMS] _loadVesselRowsFromApi warning:', err);
      this.setState({ vesselLoadingName: null, documentFilesLoading: false });
    }
  }

  /** Load the next page of vessel document rows without replacing prior pages. */
  public async _loadMoreDocumentVessels(): Promise<void> {
    const { documentVesselCount, documentVesselsLoadingMore, vessels } = this.state;
    if (documentVesselsLoadingMore || documentVesselCount >= vessels.length) return;

    this.setState({ documentVesselsLoadingMore: true });
    try {
      const params = new URLSearchParams({
        vessel_limit: '8',
        vessel_offset: String(documentVesselCount),
      });
      const data = await this._fetchJson(`${this._base()}/api/vessels/flat-tree?${params.toString()}`);
      if (!Array.isArray(data)) throw new Error('The next vessel page could not be loaded.');

      const incoming = this._normalize(data);
      const nextVesselNames = new Set(
        vessels.slice(documentVesselCount, documentVesselCount + 8)
          .map(v => cleanName(v.name).trim().toLowerCase())
      );

      this.setState(prev => {
        const fileRowKey = (row: FlatRow): string =>
          [
            cleanName(row.vesselName || '').trim().toLowerCase(),
            (row.subFolderPath || '').trim().toLowerCase(),
            (row.fileName || '').trim().toLowerCase(),
          ].join('||');
        const incomingFileKeys = new Set(
          incoming
            .filter(r => Boolean(r.fileName))
            .map(fileRowKey)
        );
        const preservedExistingFiles = prev.rows.filter(row =>
          nextVesselNames.has(cleanName(row.vesselName).trim().toLowerCase()) &&
          Boolean(row.fileName) &&
          !incomingFileKeys.has(fileRowKey(row))
        );
        // A vessel can be selected independently while this request is in
        // flight. Replace rows only for this page and retain all others.
        const retained = prev.rows.filter(row =>
          !nextVesselNames.has(cleanName(row.vesselName).trim().toLowerCase())
        );
        return {
          rows: [...retained, ...incoming, ...preservedExistingFiles],
          documentVesselCount: Math.min(prev.documentVesselCount + 8, prev.vessels.length),
          documentVesselsLoadingMore: false,
        };
      }, () => {
        // Use spoFolderMap-based refresh (fast) instead of full Graph tree walk.
        // Same live-library-scope gate as _loadData's Step 2 — see there.
        const isLiveLibraryScope = this.state.docScopeType === 'sites' ||
          this.state.docScopeType === 'shared_docs' || this.state.docScopeType === 'documents';
        if (!isLiveLibraryScope) {
          this._lastRefreshedRowsKey = '';
          this._refreshFilesFromBackendRows();
        }
      });
    } catch (err) {
      console.warn('[VesselDMS] failed to load the next vessel page:', err);
      this.setState({ documentVesselsLoadingMore: false });
    }
  }

  /** Merge existing SharePoint files for the vessels currently shown in Documents. */
  public _refreshDocumentVesselFiles(vesselNames: string[], page: number = 0): void {
    if (vesselNames.some(Boolean)) this._refreshFilesFromBackendRows(page);
  }

  /**
   * Refresh files for all rows that have a breadcrumb path (subFolderPath)
   * using path-based Graph lookup. This is the primary path for showing
   * existing files on initial load and works regardless of whether the
   * backend drive_item_ids are stale or from a different drive.
   */
  /** Run-once guard: track the last rows snapshot we refreshed from, so we
   * don't fire Graph calls repeatedly if the row set hasn't changed. */
  private _lastRefreshedRowsKey: string = '';
  private _fileFolderRefreshRequestedAt: Map<string, number> = new Map();

  public _refreshFilesFromBackendRows(page: number = 0): void {
    if (this._isUnmounted) return;
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    const { rows, documentVesselCount, vessels, vesselFilter } = this.state;
    const isSpecificVessel = vesselFilter && vesselFilter !== 'all';
    const visibleVesselNames = new Set(
      isSpecificVessel
        ? [cleanName(vesselFilter).trim().toLowerCase()]
        : vessels.slice(0, documentVesselCount).map(v => cleanName(v.name).trim().toLowerCase())
    );

    // Collect unique subFolderPaths + uploadFolderIds for visible vessels.
    const seen = new Set<string>();
    const toRefresh: Array<{ subFolderPath: string; groupKey: string; uploadFolderId: string }> = [];
    const LIST_PAGE_SIZE = 20;
    const pageStart = Math.max(0, page) * LIST_PAGE_SIZE;
    const orderedRows = [...rows].sort((a, b) =>
      (b.fileUploadedAt || 0) - (a.fileUploadedAt || 0)
    );
    const pageRows = orderedRows.slice(pageStart, pageStart + LIST_PAGE_SIZE);
    for (const row of pageRows) {
      if (!isSpecificVessel && !visibleVesselNames.has(cleanName(row.vesselName).trim().toLowerCase())) continue;
      if (isSpecificVessel && cleanName(row.vesselName).trim().toLowerCase() !== cleanName(vesselFilter).trim().toLowerCase()) continue;
      if (!row.subFolderPath || seen.has(row.subFolderPath)) continue;
      seen.add(row.subFolderPath);
      toRefresh.push({ subFolderPath: row.subFolderPath, groupKey: row.groupKey, uploadFolderId: row.uploadFolderId });
    }

    if (toRefresh.length === 0) return;

    // Deduplicate runs: skip if the exact same set of folders was already refreshed
    const rowsKey = toRefresh.map(r => r.subFolderPath).sort().join('|');
    if (rowsKey === this._lastRefreshedRowsKey) return;
    this._lastRefreshedRowsKey = rowsKey;

    const BATCH = 2;
    const runBatch = async (items: typeof toRefresh): Promise<void> => {
      for (let i = 0; i < items.length; i += BATCH) {
        if (this._isUnmounted) return;
        // Collect results for the whole batch first, then apply in ONE setState
        // call to prevent concurrent setState calls from overwriting each other.
        const batchResults: Array<{
          groupKey: string;
          uploadFolderId: string;
          parsedFiles: Array<{ name: string; size: string; date: string; pending: boolean; id: string; uploadedAt?: number }>;
        }> = [];

        await Promise.all(
          items.slice(i, i + BATCH).map(async item => {
            let fileItems: any[] = [];

            // Strategy 1: Resolve folder via live spoFolderMap (delta-synced active drive).
            const liveFolderId = this._getLiveSharePointFolderId(item.subFolderPath);
            const refreshFolderKey = `${siteId}:${driveId}:${liveFolderId || item.uploadFolderId || item.subFolderPath}`;
            if (this._fileFolderRefreshRequestedAt.has(refreshFolderKey)) return;
            this._fileFolderRefreshRequestedAt.set(refreshFolderKey, Date.now());
            if (liveFolderId) {
              const node = this.state.spoFolderMap.get(liveFolderId);
              const memFiles = (node?.children || []).filter(c => !c.isFolder);
              try {
                const url = `/sites/${siteId}/drives/${driveId}/items/${liveFolderId}/children?$select=id,name,size,createdDateTime,lastModifiedDateTime,file&$top=200`;
                const result: any = await this._getGraphChildrenResponse(url);
                const hits: any[] = (result?.value ?? []).filter((itm: any) => !!itm.file);
                if (hits.length > 0) { fileItems = hits; }
              } catch {
                fileItems = memFiles.map(c => ({
                  id: c.id,
                  name: c.name,
                  size: c.size,
                  createdDateTime: c.createdDateTime,
                  lastModifiedDateTime: c.lastModifiedDateTime,
                  file: {},
                }));
              }
            }

            // Strategy 2: uploadFolderId confirmed in active drive spoFolderMap (fallback).
            if (fileItems.length === 0 && item.uploadFolderId &&
                this.state.spoFolderMap.has(item.uploadFolderId)) {
              try {
                const url = `/sites/${siteId}/drives/${driveId}/items/${item.uploadFolderId}/children?$select=id,name,size,createdDateTime,lastModifiedDateTime,file&$top=200`;
                const result: any = await this._getGraphChildrenResponse(url);
                const hits: any[] = (result?.value ?? []).filter((itm: any) => !!itm.file);
                if (hits.length > 0) { fileItems = hits; }
              } catch {}
            }

            // Strategy 3: Direct Graph path lookup (only if folder map is empty / before baseline)
            if (fileItems.length === 0 && item.subFolderPath && this.state.spoFolderMap.size === 0) {
              const cleanPath = this._sharePointFolderPath(item.subFolderPath, '');
              if (cleanPath) {
                try {
                  const encoded = cleanPath.split('/').map(s => encodeURIComponent(s)).join('/');
                  const url = `/sites/${siteId}/drives/${driveId}/root:/${encoded}:/children?$select=id,name,size,createdDateTime,lastModifiedDateTime,file&$top=200`;
                  const result: any = await this._getGraphChildrenResponse(url);
                  const hits: any[] = (result?.value ?? []).filter((itm: any) => !!itm.file);
                  if (hits.length > 0) {
                    fileItems = hits;
                  }
                } catch {}
              }
            }

            if (fileItems.length === 0) return;

            batchResults.push({
              groupKey: item.groupKey,
              uploadFolderId: item.uploadFolderId,
              parsedFiles: fileItems.map((f: any) => ({
                name: f.name,
                size: f.size ? `${(f.size / 1024).toFixed(1)} KB` : '—',
                date: f.createdDateTime ? new Date(f.createdDateTime).toLocaleString([], { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : (f.lastModifiedDateTime || 'Today'),
                pending: false,
                id: f.id,
                uploadedAt: f.createdDateTime ? Date.parse(f.createdDateTime) : (f.lastModifiedDateTime ? Date.parse(f.lastModifiedDateTime) : undefined),
              })),
            });
          })
        );

        if (this._isUnmounted) return;

        // Apply ALL batch results in ONE setState to avoid race conditions.
        if (batchResults.length > 0) {
          this.setState(prev => {
            let updatedRows = prev.rows;
            const updatedByFolder = { ...prev.uploadedFilesByFolder };

            for (const { groupKey, uploadFolderId, parsedFiles } of batchResults) {
              const baseRow = updatedRows.find(r => r.groupKey === groupKey);
              if (!baseRow) continue;

              // Preserve any existing rows that have files
              const existingWithFile = updatedRows.filter(r => r.groupKey === groupKey && Boolean(r.fileName));
              const otherRows = updatedRows.filter(r => r.groupKey !== groupKey);
              const mappedRows = parsedFiles.map((f, idx) => ({
                ...baseRow,
                srNo: idx === 0 ? baseRow.srNo : `${baseRow.srNo}.${idx + 1}`,
                fileName: f.name,
                fileId: f.id,
                filePending: false,
                fileUploadedAt: f.uploadedAt,
                fileSize: f.size !== '—' ? f.size : undefined,
              }));

              const preservedRows: FlatRow[] = [];
              existingWithFile.forEach(er => {
                if (!parsedFiles.some(pf => pf.name.toLowerCase() === (er.fileName || '').toLowerCase())) {
                  preservedRows.push(er);
                }
              });

              updatedRows = [...otherRows, ...mappedRows, ...preservedRows];

              // Merge into uploadedFilesByFolder
              const curFiles = updatedByFolder[groupKey] || [];
              const mergedUploads = [...parsedFiles];
              curFiles.forEach((cf: any) => {
                if (cf?.name && !mergedUploads.some(mf => mf.name.toLowerCase() === cf.name.toLowerCase())) {
                  mergedUploads.push(cf);
                }
              });

              updatedByFolder[groupKey] = mergedUploads;
              if (uploadFolderId) updatedByFolder[uploadFolderId] = mergedUploads;
            }

            return { rows: updatedRows, uploadedFilesByFolder: updatedByFolder };
          });
        }
      }
    };
    void runBatch(toRefresh);
  }

  /**
   * Merge folders and files already present in SharePoint into the backend list rows.
   * This deliberately walks both the current Vessels/{vessel} layout and the
   * legacy Main Folder/{vessel} layout handled by _flattenVesselViaGraph.
   */
  public async _mergeLiveSharePointFiles(vesselNames: string[]): Promise<void> {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId || vesselNames.length === 0) {
      if (this.state.documentFilesLoading) this.setState({ documentFilesLoading: false });
      return;
    }

    const mergeKey = vesselNames.filter(Boolean).map(name => name.trim().toLowerCase()).sort().join('|');
    if (this._liveSharePointMerges.has(mergeKey)) return;
    this._liveSharePointMerges.add(mergeKey);
    this.setState({ documentFilesLoading: true });
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), 20000);
    try {
      const signal = controller.signal;
      const liveFolderRows: FlatRow[] = [];
      await this._mapLimit(vesselNames.filter(Boolean), 4, async vesselName => {
        const rows = await this._flattenVesselViaGraph(vesselName, signal, () => undefined).catch(() => []);
        // _walkGraphFolder emits one row for every uploadable leaf, even when
        // it has no files. Keeping those empty rows is essential: their
        // uploadFolderId is the actual SPO location used for later uploads.
        liveFolderRows.push(...rows.filter(row => Boolean(row.uploadFolderId)));
      });
      if (liveFolderRows.length === 0) {
        this.setState({ documentFilesLoading: false });
        return;
      }

      const keyFor = (row: Pick<FlatRow, 'vesselName' | 'group' | 'category' | 'subCategory' | 'subFolderPath'>): string =>
        [row.vesselName, row.group, row.category, row.subCategory || row.category, row.subFolderPath]
          .map(value => cleanName(value || '').trim().toLowerCase())
          .join('||');

      this.setState(previous => {
        const baseRows = new Map<string, FlatRow>();
        previous.rows.forEach(row => {
          const key = keyFor(row);
          if (!baseRows.has(key)) baseRows.set(key, row);
        });

        const mergedFolderRows: FlatRow[] = [];
        const affectedKeys = new Set<string>();
        const liveRowsByKey = new Map<string, FlatRow[]>();
        liveFolderRows.forEach(row => {
          const key = keyFor(row);
          const rows = liveRowsByKey.get(key) || [];
          rows.push(row);
          liveRowsByKey.set(key, rows);
        });

        liveRowsByKey.forEach((folderRows, semanticKey) => {
          const liveFolder = folderRows[0];
          const baseRow = baseRows.get(semanticKey);
          // Preserve backend metadata/group key where available, but always use
          // the path obtained by listing the active SharePoint drive.
          const rowBase: FlatRow = baseRow
            ? {
                ...baseRow,
                uploadFolderId: liveFolder.uploadFolderId,
                monthDriven: liveFolder.monthDriven,
                canUpload: liveFolder.canUpload,
                filePending: false,
              }
            : liveFolder;
          const files = folderRows.filter(row => Boolean(row.fileName));
          affectedKeys.add(semanticKey);

          if (files.length === 0) {
            // Preserve any file already in state for this folder (e.g. from DB or previous upload)
            const existingFilesInState = previous.rows.filter(
              row => keyFor(row) === semanticKey && Boolean(row.fileName)
            );
            if (existingFilesInState.length > 0) {
              existingFilesInState.forEach(row => mergedFolderRows.push({
                ...row,
                uploadFolderId: liveFolder.uploadFolderId || row.uploadFolderId,
                monthDriven: liveFolder.monthDriven ?? row.monthDriven,
                canUpload: liveFolder.canUpload ?? row.canUpload,
              }));
            } else if (baseRow?.fileName) {
              mergedFolderRows.push({
                ...baseRow,
                uploadFolderId: liveFolder.uploadFolderId || baseRow.uploadFolderId,
              });
            } else {
              mergedFolderRows.push({ ...rowBase, fileName: null, fileId: null, filePending: false });
            }
            return;
          }

          const seenFiles = new Set<string>();
          // First preserve any existing files from state
          const existingFilesInState = previous.rows.filter(
            row => keyFor(row) === semanticKey && Boolean(row.fileName)
          );
          existingFilesInState.forEach(row => {
            const fileKey = (row.fileName || '').toLowerCase();
            if (!seenFiles.has(fileKey)) {
              seenFiles.add(fileKey);
              mergedFolderRows.push({
                ...row,
                uploadFolderId: liveFolder.uploadFolderId || row.uploadFolderId,
              });
            }
          });

          files.forEach(fileRow => {
            const fileKey = (fileRow.fileName || '').toLowerCase();
            if (!seenFiles.has(fileKey)) {
              seenFiles.add(fileKey);
              mergedFolderRows.push({
                ...rowBase,
                fileName: fileRow.fileName,
                fileId: fileRow.fileId,
                filePending: false,
              });
            }
          });
        });

        // Replace rows for folders confirmed by the live SPO walk, including
        // empty folders, while preserving backend-only template rows.
        const retained = previous.rows.filter(row => !affectedKeys.has(keyFor(row)));
        const updatedByFolder = { ...previous.uploadedFilesByFolder };

        mergedFolderRows.forEach(row => {
          if (row.fileName) {
            const newF = {
              name: row.fileName,
              size: row.fileSize || '—',
              date: row.fileUploadedAt ? new Date(row.fileUploadedAt).toLocaleString() : 'Today',
              pending: false,
              id: row.fileId || row.fileName,
              uploadedAt: row.fileUploadedAt,
            };
            const addToList = (key: string | undefined | null): void => {
              if (!key) return;
              const list = updatedByFolder[key] || [];
              if (!list.some((f: any) => f.name.toLowerCase() === row.fileName!.toLowerCase())) {
                updatedByFolder[key] = [...list, newF];
              }
            };
            addToList(row.groupKey);
            addToList(row.uploadFolderId);
            addToList(row.subFolderPath);
            if (row.subFolderPath) addToList(row.subFolderPath.trim().toLowerCase());
            addToList(row.subCategory);
            if (row.subCategory) addToList(row.subCategory.trim().toLowerCase());
            addToList(row.category);
            if (row.category) addToList(row.category.trim().toLowerCase());
            if (row.vesselName && row.subCategory) {
              addToList(`${row.vesselName} > ${row.subCategory}`.toLowerCase());
              addToList(`${row.vesselName} > ${row.group} > ${row.subCategory}`.toLowerCase());
              addToList(`${row.group} > ${row.vesselName} > ${row.subCategory}`.toLowerCase());
            }
          }
        });

        return { rows: [...retained, ...mergedFolderRows], uploadedFilesByFolder: updatedByFolder, documentFilesLoading: false };
      });
    } catch (err) {
      console.warn('[VesselDMS] _mergeLiveSharePointFiles warning:', err);
      this.setState({ documentFilesLoading: false });
    } finally {
      window.clearTimeout(timeoutId);
      this._liveSharePointMerges.delete(mergeKey);
      if (this.state.documentFilesLoading) this.setState({ documentFilesLoading: false });
    }
  }

  /** Open the exact SharePoint folder represented by a Documents list row. */
  /** SharePoint Online URL of the folder that holds a List view row's files,
   * resolved from folder listings already loaded into _siteFolderItemsCache
   * (each Graph item there carries webUrl + parentReference). Returns '' when
   * nothing cached identifies it. */
  private _cachedFolderUrlForRow(row: GroupedRow): string {
    const folderId = (row.uploadFolderId || '').trim();
    const fileIds = new Set((row.files || []).map(f => f.id).filter(Boolean));
    const isOfficeViewerUrl = (url: string): boolean => /\/_layouts\//i.test(url);
    let folderUrl = '';
    let parentIdFromFile = '';
    let fileBasedUrl = '';
    this._siteFolderItemsCache.forEach(entry => {
      if (folderUrl || !entry || !Array.isArray(entry.items)) return;
      entry.items.forEach((item: any) => {
        if (folderUrl || !item?.id) return;
        const url: string = item.webUrl || item.web_url || '';
        if (item.folder && folderId && item.id === folderId && url) {
          folderUrl = url;
        } else if (!item.folder && fileIds.has(item.id)) {
          if (!parentIdFromFile && item.parentReference?.id) parentIdFromFile = item.parentReference.id;
          // A non-Office file's webUrl is ".../Folder/file.pdf" — its folder
          // is that URL minus the last segment. Office files open through
          // /_layouts/15/Doc.aspx, which carries no folder path.
          if (!fileBasedUrl && url && !isOfficeViewerUrl(url) && url.lastIndexOf('/') > 8) {
            fileBasedUrl = url.slice(0, url.lastIndexOf('/'));
          }
        }
      });
    });
    if (!folderUrl && parentIdFromFile && parentIdFromFile !== folderId) {
      this._siteFolderItemsCache.forEach(entry => {
        if (folderUrl || !entry || !Array.isArray(entry.items)) return;
        const hit = entry.items.find((item: any) => item?.folder && item.id === parentIdFromFile);
        if (hit) folderUrl = hit.webUrl || hit.web_url || '';
      });
    }
    return folderUrl || fileBasedUrl;
  }

  public async _openSharePointFolder(row: GroupedRow): Promise<void> {
    // Fast path: the folder's SharePoint URL is usually already known from
    // the folder listings this row was built from — open it directly.
    const cachedFolderUrl = this._cachedFolderUrlForRow(row);
    if (cachedFolderUrl) {
      window.open(cachedFolderUrl, '_blank');
      return;
    }
    // Open the new tab synchronously (inside the user-gesture frame) so the
    // browser does not treat the later navigation as a popup.  We deliberately
    // omit 'noopener' here because modern Chrome/Edge return null when noopener
    // is used, which means target is null and the async redirect below would be
    // blocked by the popup blocker – leaving the user with an about:blank tab.
    // The destination is always the same SharePoint tenant, so the window.opener
    // reference in the new tab is not a security concern.
    const target = window.open('about:blank', '_blank');
    const { graphClient } = this.props;

    // Resolve the effective site (siteId, driveId, siteUrl) from state.documentSites.
    // The row may carry siteKey (set by DocumentsPage), otherwise fall back to
    // activeDocumentSite. Only default to the web-part props when the row
    // genuinely belongs to the primary Communication site.
    const rowSiteKey: string | undefined = (row as any).siteKey;
    const effectiveSiteKey = rowSiteKey || this.state.activeDocumentSite || '';
    const matchedSite = effectiveSiteKey
      ? (this.state.documentSites || []).find(
          s => s.site_key.toLowerCase() === effectiveSiteKey.toLowerCase()
        )
      : null;

    const siteId   = (matchedSite?.site_id  && matchedSite.site_id !== matchedSite.site_key)
                       ? matchedSite.site_id  : this.props.siteId;
    const driveId  = matchedSite?.drive_id  || this.props.driveId;
    const siteUrl  = matchedSite?.web_url   || this.props.siteUrl || '';

    const normalisePath = (path: string): string =>
      (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
    const sharePointPath = this._sharePointFolderPath(row.subFolderPath, '');
    const navigationId = `nav_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    const logNavigation = (event: string, details: Record<string, any> = {}): void => {
      const payload = { navigation_id: navigationId, event, ui_path: row.subFolderPath, sharepoint_path: sharePointPath, details };
      console.info(`[VesselDMS] folder navigation ${event}`, payload);
      void fetch(`${this._base()}/api/diagnostics/sharepoint-navigation`, {
        method: 'POST',
        headers: { ...this._headers(), 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        keepalive: true,
      }).catch(() => undefined);
    };
    console.info('[VesselDMS] folder navigation start', {
      uiPath: row.subFolderPath,
      sharePointPath,
      effectiveSiteKey,
      siteUrl,
      siteId,
      driveId,
    });
    logNavigation('start', { siteUrl, siteId, driveId, effectiveSiteKey });
    const expectedPath = normalisePath(sharePointPath);
    let liveFolderPath: string | undefined;
    let resolvedDrivePath: string | undefined;

    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder) continue;
      const path = normalisePath(node.serverRelativePath);
      if (path === expectedPath || (expectedPath && path.endsWith(`/${expectedPath}`))) {
        liveFolderPath = node.serverRelativePath;
        break;
      }
    }

    try {
      if (!graphClient || !siteId || !driveId) {
        throw new Error('SharePoint connection is not available for this web part.');
      }

      // Resolve by path in the active browser drive. Backend folder IDs and
      // delta item IDs can belong to an older drive even when their paths match.
      let item: any;
      // First, the row's own folder/file item id on the row's own site/drive
      // (SharePoint Sites rows carry Graph ids from that site's drive).
      const itemIdCandidates = [row.uploadFolderId, ...(row.files || []).map(f => f.id)]
        .map(id => String(id || '').trim())
        .filter(id => id && !id.includes('/') && !/^(sf_|file_|category_|site:|drive:)/i.test(id) && !/^\d+$/.test(id));
      for (const candidateId of Array.from(new Set(itemIdCandidates))) {
        try {
          const byId: any = await graphClient
            .api(`/sites/${siteId}/drives/${driveId}/items/${encodeURIComponent(candidateId)}?$select=id,folder,file,webUrl,parentReference`)
            .get();
          if (byId?.folder && byId.webUrl) {
            item = byId;
          } else if (byId?.parentReference?.id) {
            const parent: any = await graphClient
              .api(`/sites/${siteId}/drives/${driveId}/items/${encodeURIComponent(byId.parentReference.id)}?$select=id,folder,webUrl`)
              .get();
            if (parent?.webUrl) item = parent;
          }
          if (item) {
            logNavigation('graph_item_id_success', { itemId: candidateId, webUrl: item.webUrl });
            break;
          }
        } catch (idError) {
          logNavigation('graph_item_id_failed', { itemId: candidateId, error: String(idError) });
        }
      }
      const paths = [
        liveFolderPath,
        sharePointPath,
        sharePointPath.replace(/\/Vessels\//i, '/'),
        `Vessels/${sharePointPath}`,
        `Vessels/Specific Vessels/${sharePointPath}`,
      ];
      let lastError: any;
      for (const candidatePath of Array.from(new Set(paths.filter(Boolean))) as string[]) {
        if (item) break;
        try {
          const driveRelativePath = candidatePath.replace(/^\/+/, '');
          const encodedPath = driveRelativePath.split('/').map(part => encodeURIComponent(part)).join('/');
          const candidate: any = await graphClient
            .api(`/sites/${siteId}/drives/${driveId}/root:/${encodedPath}?$select=id,folder,webUrl`)
            .get();
          if (candidate?.folder && candidate?.webUrl) {
            item = candidate;
            resolvedDrivePath = driveRelativePath;
            console.info('[VesselDMS] folder navigation success via Graph path', {
              path: driveRelativePath,
              folderId: candidate.id,
              webUrl: candidate.webUrl,
            });
            logNavigation('graph_path_success', { path: driveRelativePath, folderId: candidate.id, webUrl: candidate.webUrl });
            break;
          }
        } catch (pathError) {
          lastError = pathError;
          console.warn('[VesselDMS] folder navigation Graph path failed', {
            path: candidatePath,
            error: pathError,
          });
          logNavigation('graph_path_failed', { path: candidatePath, error: String(pathError) });
        }
      }
      if (!item) {
        const pathParts = sharePointPath.split('/').filter(Boolean);
        const leafName = pathParts[pathParts.length - 1] || '';
        const vesselName = pathParts[0] || '';
        const mainName = pathParts[1] || '';
        try {
          const searchUrl = `/sites/${siteId}/drives/${driveId}/root/search(q='${encodeURIComponent(leafName)}')?$select=id,name,folder,parentReference,webUrl&$top=200`;
          const searchResult: any = await graphClient.api(searchUrl).get();
          const candidates = (searchResult?.value || []).filter((candidate: any) => {
            if (!candidate?.folder || !candidate?.webUrl) return false;
            if ((candidate.name || '').toLocaleLowerCase() !== leafName.toLocaleLowerCase()) return false;
            const parentPath = decodeURIComponent(candidate.parentReference?.path || '').toLocaleLowerCase();
            return parentPath.includes(vesselName.toLocaleLowerCase()) && parentPath.includes(mainName.toLocaleLowerCase());
          });
          if (candidates.length > 0) {
            item = candidates[0];
            const parentPath = decodeURIComponent(item.parentReference?.path || '');
            const rootIndex = parentPath.indexOf('/root:');
            if (rootIndex >= 0) {
              resolvedDrivePath = `${parentPath.slice(rootIndex + 6).replace(/^\/+/, '')}/${item.name}`;
            }
          }
          if (item) {
            console.info('[VesselDMS] folder navigation success via Graph search', {
              folderId: item.id,
              webUrl: item.webUrl,
            });
          }
          if (item) logNavigation('graph_search_success', { folderId: item.id, webUrl: item.webUrl });
        } catch (searchError) {
          lastError = searchError;
          console.warn('[VesselDMS] folder navigation Graph search failed', searchError);
          logNavigation('graph_search_failed', { error: String(searchError) });
        }
      }
      if (!item && siteUrl) {
        const site = new URL(siteUrl);
        const sitePath = site.pathname.replace(/\/$/, '');
        const restPaths = [
          `${sitePath}/Shared Documents/Vessels/Specific Vessels/${sharePointPath}`,
          `${sitePath}/Shared Documents/Vessels/${sharePointPath}`,
          `${sitePath}/Shared Documents/Vessel Management/${sharePointPath}`,
          `${sitePath}/Shared Documents/${sharePointPath}`,
        ];
        for (const serverRelativePath of Array.from(new Set(restPaths))) {
          try {
            const escapedPath = serverRelativePath.replace(/'/g, "''");
            const response = await fetch(
              `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeURIComponent(escapedPath)}')?$select=ServerRelativeUrl`,
              { headers: { Accept: 'application/json;odata=nometadata' }, credentials: 'include' },
            );
            if (response.ok) {
              const folder = await response.json();
              const actualPath = folder?.ServerRelativeUrl || serverRelativePath;
              // SharePoint AllItems.aspx ?id= needs the server-relative path with
              // literal slashes – encodeURIComponent would turn '/' into '%2F' and
              // produce a blank page.  Encode only non-slash segments instead.
              const encodedActualPath = actualPath.split('/').map((s: string) => encodeURIComponent(s)).join('/');
              item = {
                webUrl: `${site.origin}${encodedActualPath}`,
              };
              resolvedDrivePath = actualPath.replace(`${sitePath}/Shared Documents/`, '');
              console.info('[VesselDMS] folder navigation success via SharePoint REST', {
                serverRelativePath: actualPath,
                webUrl: item.webUrl,
              });
              logNavigation('sharepoint_rest_success', { serverRelativePath: actualPath, webUrl: item.webUrl });
              break;
            }
            console.warn('[VesselDMS] folder navigation SharePoint REST path failed', {
              serverRelativePath,
              status: response.status,
            });
            logNavigation('sharepoint_rest_failed', { serverRelativePath, status: response.status });
          } catch (restError) {
            lastError = restError;
            console.warn('[VesselDMS] folder navigation SharePoint REST request failed', restError);
            logNavigation('sharepoint_rest_error', { error: String(restError) });
          }
        }
      }
      if (!item) {
        const fallbackBase = siteUrl || this.props.siteUrl;
        if (fallbackBase) {
          try {
            const site = new URL(fallbackBase);
            const sitePath = site.pathname.replace(/\/$/, '');
            if (target) {
              const fallbackUrl = `${site.origin}${sitePath}/Shared%20Documents`;
              console.warn('[VesselDMS] folder navigation fallback to Documents root', {
                url: fallbackUrl,
                error: lastError,
              });
              logNavigation('fallback_documents_root', { url: fallbackUrl, error: String(lastError || '') });
              target.location.href = fallbackUrl;
            }
          } catch { /* malformed URL — fall through */ }
          return;
        }
        throw lastError || new Error('The selected SharePoint folder was not found.');
      }
      if (!item?.webUrl) throw new Error('SharePoint did not return a folder link.');
      // item.webUrl from Graph / SharePoint REST is the canonical, direct URL to open
      // the folder in the modern document library view without fragile AllItems query params.
      const redirectUrl = item.webUrl;
      console.info('[VesselDMS] folder navigation redirecting', { webUrl: redirectUrl, resolvedDrivePath });
      logNavigation('redirecting', { url: redirectUrl, resolvedDrivePath });
      if (target) {
        target.location.href = redirectUrl;
        target.focus();
      } else {
        // target is null only when popup blockers prevented the initial open.
        // Try a direct open – the user will need to allow popups for this site.
        window.open(redirectUrl, '_blank');
      }
    } catch (err: any) {
      if (target) target.close();
      console.error('[VesselDMS] folder navigation failed', {
        uiPath: row.subFolderPath,
        sharePointPath,
        error: err,
      });
      logNavigation('failed', { error: String(err) });
      alert(err?.message || 'Unable to open this SharePoint folder.');
    }
  }

  /** Open a file from the active SharePoint Documents drive. */
  public async _openDocumentFile(fileId: string, fileName: string, folderPath?: string): Promise<void> {
    const { graphClient, siteId, driveId } = this.props;

    // 1. If Graph client is available, attempt to resolve the SharePoint webUrl
    if (graphClient && siteId && driveId) {
      // A. Try direct lookup by Graph Drive Item ID if valid
      if (fileId && !fileId.startsWith('file_') && !/^\d{13,}$/.test(fileId) && fileId !== fileName) {
        try {
          const item: any = await graphClient
            .api(`/sites/${siteId}/drives/${driveId}/items/${encodeURIComponent(fileId)}?$select=id,name,webUrl,file`)
            .get();
          if (item?.webUrl) {
            window.open(item.webUrl, '_blank');
            return;
          }
        } catch {}
      }

      // B. Try path-based lookup if folderPath and fileName are provided
      if (folderPath && fileName) {
        const cleanDrivePath = this._sharePointFolderPath(folderPath, folderPath);
        const candidatePaths = [
          `${cleanDrivePath}/${fileName}`,
          `Kaizen - Knowledge Bank/${cleanDrivePath.replace(/^kaizen\s*-\s*knowledge\s*bank\//i, '')}/${fileName}`,
          `Vessels/${cleanDrivePath}/${fileName}`,
          `Vessels/Specific Vessels/${cleanDrivePath}/${fileName}`,
        ];
        for (const cand of Array.from(new Set(candidatePaths))) {
          try {
            const encodedPath = cand.split('/').map(part => encodeURIComponent(part)).join('/');
            const item: any = await graphClient
              .api(`/sites/${siteId}/drives/${driveId}/root:/${encodedPath}?$select=id,name,webUrl,file`)
              .get();
            if (item?.webUrl) {
              window.open(item.webUrl, '_blank');
              return;
            }
          } catch {}
        }
      }

      // C. Try Graph search by filename
      if (fileName) {
        try {
          const searchResult: any = await graphClient
            .api(`/sites/${siteId}/drives/${driveId}/root/search(q='${encodeURIComponent(fileName)}')?$select=id,name,webUrl,file&$top=20`)
            .get();
          const match = (searchResult?.value || []).find((f: any) => (f.name || '').toLowerCase() === fileName.toLowerCase());
          if (match?.webUrl) {
            window.open(match.webUrl, '_blank');
            return;
          }
        } catch {}
      }
    }

    // 2. Fallback to backend download content endpoint
    if (fileId && !fileId.startsWith('file_') && !/^\d{13,}$/.test(fileId)) {
      window.open(`${this._base()}/api/files/${encodeURIComponent(fileId)}/content`, '_blank');
      return;
    }

    // 3. Fallback: if we have siteUrl, open document library
    if (this.props.siteUrl) {
      window.open(`${this.props.siteUrl}/Shared Documents`, '_blank');
    }
  }

  /** Resolve a UI breadcrumb to the matching folder in the active SPO drive. */
  public _getLiveSharePointFolderId(subFolderPath: string): string | null {
    if (!subFolderPath) return null;
    const parts = subFolderPath.split('>').map(p => p.trim()).filter(Boolean);
    if (parts.length === 0) return null;

    const normalisePath = (path: string): string =>
      (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
    const expectedPath = normalisePath(this._sharePointFolderPath(subFolderPath, ''));

    const cleanKey = (s: string) => {
      const lower = (s || '').toLowerCase().trim();
      if (/^common for all (ships|vessels)$/i.test(lower) || lower === 'common') return 'commonforallships';
      if (/^kaizen(\s*-\s*knowledge bank)?$/i.test(lower) || lower === 'knowledge bank') return 'kaizenknowledgebank';
      return lower
        .replace(/^folder-\d+\s+/i, '')
        .replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '')
        .replace(/[^\w\d]/g, '');
    };

    const cleanLeaf = cleanKey(parts[parts.length - 1]);
    const cleanVessel = cleanKey(parts[0]);
    const cleanGroup = parts.length > 1 ? cleanKey(parts[1]) : '';
    const cleanParent = parts.length >= 2 ? cleanKey(parts[parts.length - 2]) : '';

    const legacyPath = parts.length >= 3
      ? normalisePath([parts[1], parts[0], ...parts.slice(2)].join('/'))
      : '';

    const isCommon = parts[0] && /^common for all (vessels|ships)$/i.test(parts[0]);
    const isKaizen = parts[0] && /^kaizen/i.test(parts[0]);
    const cleanPaths: string[] = [];
    if (isCommon) {
      const rest = parts.slice(1).join('/');
      cleanPaths.push(normalisePath(`Vessels/Common for all ships/${rest}`));
      cleanPaths.push(normalisePath(`Common for all ships/${rest}`));
      cleanPaths.push(normalisePath(`Vessels/Common for all vessels/${rest}`));
      cleanPaths.push(normalisePath(`Common for all vessels/${rest}`));
    } else if (isKaizen) {
      const rest = parts.slice(1).join('/');
      cleanPaths.push(normalisePath(`Kaizen - Knowledge Bank/${rest}`));
      cleanPaths.push(normalisePath(`Kaizen/${rest}`));
    }
    if (expectedPath) cleanPaths.push(expectedPath);
    if (legacyPath) cleanPaths.push(legacyPath);

    // First pass: exact or suffix path matches (fastest and most specific)
    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder) continue;
      const livePath = normalisePath(node.serverRelativePath);
      for (const p of cleanPaths) {
        if (p && (livePath === p || livePath.endsWith(`/${p}`))) return node.id;
      }
    }

    // Second pass: segment token matching (handles custom folder prefixes and structure variations)
    const minSegments = parts.length;
    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder) continue;
      const nodeSegments = (node.serverRelativePath || '').split('/').filter(Boolean).map(cleanKey);
      if (nodeSegments.length < minSegments) continue;
      const lastSeg = nodeSegments[nodeSegments.length - 1];

      if (lastSeg === cleanLeaf || (lastSeg && cleanLeaf && (lastSeg.includes(cleanLeaf) || cleanLeaf.includes(lastSeg)))) {
        const parentSeg = nodeSegments[nodeSegments.length - 2] || '';
        const parentMatches = !cleanParent || parentSeg === cleanParent ||
          (parentSeg && (parentSeg.includes(cleanParent) || cleanParent.includes(parentSeg)));
        if (!parentMatches) continue;
        const hasVessel = nodeSegments.some(seg => seg === cleanVessel || (cleanVessel && seg.includes(cleanVessel)));
        if (hasVessel) {
          if (!cleanGroup) return node.id;
          const hasGroup = nodeSegments.some(seg => seg === cleanGroup || (cleanGroup && (seg.includes(cleanGroup) || cleanGroup.includes(seg))));
          if (hasGroup) return node.id;
        }
      }
    }

    const liveFolders = this.state.documentLiveFolders || [];
    const livePathMatches = liveFolders.filter(folder => {
      const path = (folder.path || '').toLowerCase().replace(/\\/g, '/');
      return cleanPaths.some((candidate: string) => {
        const normalized = candidate.toLowerCase().replace(/\\/g, '/');
        return normalized && (path === normalized || path.endsWith(`/${normalized}`));
      });
    });
    if (livePathMatches.length > 0) return livePathMatches[0].id;

    return null;
  }

  /**
   * Upload a file to a SharePoint folder via Graph API (direct SPO upload) with REST API fallback.
   * Prevents FastAPI 404 path routing errors when folder path strings containing slashes are used.
   */
  public async _uploadFileToFolder(
    uploadFolderId: string,
    subFolderPath: string,
    vesselName: string,
    file: File,
    monthDriven: boolean = false,
    targetSiteId?: string,
    targetDriveId?: string,
    baseSubFolderPath?: string,
  ): Promise<{ fileId: string | null; statusPending: boolean; folderId: string | null; isGraphUpload: boolean; webUrl: string | null }> {
    const base = this._base();
    if (!base) return { fileId: null, statusPending: false, folderId: null, isGraphUpload: false, webUrl: null };

    const effectiveSiteId = targetSiteId || this.props.siteId;
    const effectiveDriveId = targetDriveId || this.props.driveId;

    const activeMainFolder = this.state.docMainFolder || 'Technical & Crewing';
    const encodePath = (p: string): string =>
      p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

    const { graphClient } = this.props;
    const rawSegments = (subFolderPath || uploadFolderId || '')
      .split(/[>/]/)
      .map(s => s.trim())
      .filter(Boolean)
      .filter(s => !/^(sharepoint sites|sites documents|documents|shared documents|home)$/i.test(s));
    const cleanDirectPath = rawSegments.join('/');
    const cleanUploadFolder = (uploadFolderId || '').replace(/^[\s>/]+|[\s>/]+$/g, '');
    const isKaizen = vesselName === 'Kaizen - Knowledge Bank' ||
      rawSegments.some(s => s.toLowerCase() === 'kaizen - knowledge bank') ||
      cleanUploadFolder.toLowerCase().includes('kaizen');
    const isCommon = vesselName === 'Common for all vessels' ||
      vesselName === 'Common for all ships' ||
      rawSegments.some(s => /^common for all (vessels|ships)$/i.test(s)) ||
      cleanUploadFolder.toLowerCase().includes('common for all');
    let canonicalSharePointPath: string;
    if (targetSiteId) {
      canonicalSharePointPath = cleanDirectPath || 'General';
    } else if (isKaizen) {
      const kaizenParts = rawSegments.filter(s => s.toLowerCase() !== 'kaizen - knowledge bank');
      canonicalSharePointPath = ['Kaizen - Knowledge Bank', ...kaizenParts].join('/');
    } else if (isCommon) {
      const commonParts = rawSegments.filter(s =>
        !/^common for all (vessels|ships)$/i.test(s) &&
        !this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === s.toLowerCase()) &&
        !/^vessels$/i.test(s) &&
        !/^specific vessels$/i.test(s)
      );
      const mainF = rawSegments.find(s => this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === s.toLowerCase())) || activeMainFolder;
      canonicalSharePointPath = [mainF, 'Common for all ships', ...commonParts].join('/');
    } else {
      const effectiveVessel = vesselName && vesselName !== 'all' && !/^(sharepoint sites|sites documents)$/i.test(vesselName)
        ? vesselName
        : (rawSegments.find(s => !this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === s.toLowerCase()) && !/^(vessels|specific vessels|documents)$/i.test(s)) || 'newvessel37');
      const mainF = rawSegments.find(s => this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === s.toLowerCase())) || activeMainFolder;
      const subTree = rawSegments.filter(s =>
        s.toLowerCase() !== effectiveVessel.toLowerCase() &&
        !this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === s.toLowerCase()) &&
        !/^(vessels|specific vessels|documents|root|common for all vessels|common for all ships)$/i.test(s)
      );
      canonicalSharePointPath = [mainF, effectiveVessel, ...subTree].join('/');
    }
    if (graphClient && effectiveSiteId && effectiveDriveId) {
      try {
        let folder: any = null;

        // 0. Direct Graph Item ID check if uploadFolderId is a real DriveItem ID or root
        if (uploadFolderId === 'root') {
          try {
            const rootItem = await graphClient
              .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/root?$select=id,name,folder,webUrl`)
              .get();
            if (rootItem?.id && rootItem.folder) folder = rootItem;
          } catch {
            // Fall back
          }
        } else if (
          uploadFolderId &&
          !uploadFolderId.includes('/') &&
          !uploadFolderId.includes('>') &&
          !/^(sf_|dept_|lib:|category_|sites_root|site:|drive:)/i.test(uploadFolderId) &&
          !/^\d+$/.test(uploadFolderId)
        ) {
          try {
            const itemById = await graphClient
              .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${uploadFolderId}?$select=id,name,folder,webUrl`)
              .get();
            if (itemById?.id && itemById.folder) {
              folder = itemById;
              console.log(`[VesselDMS] _uploadFileToFolder: resolved folder in SPO directly by ID "${uploadFolderId}" (name="${folder.name}")`);
            }
          } catch {
            // Fall back to candidate paths
          }
        }

        // Folder uploads pass the current folder ID plus a deeper relative path.
        // Resolve that relative path beneath the current folder before uploading.
        if (folder && baseSubFolderPath && subFolderPath) {
          const splitPath = (path: string): string[] => path.split(/[>/]/).map(part => part.trim()).filter(Boolean);
          const baseSegments = splitPath(baseSubFolderPath);
          const targetSegments = splitPath(subFolderPath);
          const hasBasePrefix = baseSegments.every((segment, index) =>
            targetSegments[index]?.toLowerCase() === segment.toLowerCase()
          );
          const relativeSegments = hasBasePrefix ? targetSegments.slice(baseSegments.length) : [];

          for (const segment of relativeSegments) {
            const encodedSegment = encodeURIComponent(segment);
            let parentFolder = folder;
            let currentCreatedOrExisting: any = null;
            try {
              const existing = await graphClient
                .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}:/${encodedSegment}?$select=id,name,folder,webUrl`)
                .get();
              if (existing?.id && existing.folder) {
                folder = existing;
                currentCreatedOrExisting = existing;
              }
            } catch {
              // The relative folder does not exist yet.
            }

            if (!currentCreatedOrExisting) {
              const created = await graphClient
                .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}/children`)
                .post({
                  name: segment,
                  folder: {},
                  '@microsoft.graph.conflictBehavior': 'fail',
                });
              folder = created;
              currentCreatedOrExisting = created;
            }

            if (parentFolder?.id && currentCreatedOrExisting) {
              const pCacheKey = `${effectiveSiteId}::${effectiveDriveId}::${parentFolder.id}`;
              const pCache = this._siteFolderItemsCache.get(pCacheKey);
              const pItem = {
                id: currentCreatedOrExisting.id,
                name: segment,
                folder: { childCount: 0 },
                webUrl: currentCreatedOrExisting.webUrl || '',
                lastModifiedDateTime: new Date().toISOString(),
              };
              if (pCache) {
                if (!pCache.items.some((it: any) => (it.name || '').toLowerCase() === segment.toLowerCase())) {
                  pCache.items = [pItem, ...pCache.items];
                }
              } else {
                this._siteFolderItemsCache.set(pCacheKey, { items: [pItem], loading: false, parentPath: '' });
              }
            }
          }
        }

        // 1. Construct canonical target SharePoint folder path if not resolved by ID
        if (!folder) {
          // Direct candidate path checks in Graph
          const rawCandidatePaths = targetSiteId
            ? [
                cleanDirectPath,
                canonicalSharePointPath,
                uploadFolderId && uploadFolderId.includes('/') ? uploadFolderId.replace(/^\/+|\/+$/g, '') : '',
              ]
            : [
                cleanDirectPath,
                canonicalSharePointPath,
                uploadFolderId && uploadFolderId.includes('/') ? uploadFolderId.replace(/^\/+|\/+$/g, '') : '',
                this._sharePointFolderPath(subFolderPath, uploadFolderId).replace(/^\/+|\/+$/g, ''),
                vesselName ? `Vessels/Specific Vessels/${vesselName}/${canonicalSharePointPath}` : '',
                vesselName ? `Vessels/${vesselName}/${canonicalSharePointPath}` : '',
              ];
          const candidatePaths = Array.from(new Set(rawCandidatePaths.filter(Boolean)))
            .filter(p => !targetSiteId || !/^(sharepoint sites|sites documents)/i.test(p.trim()));

          for (const tryPath of candidatePaths) {
            try {
              const folderItem = await graphClient
                .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/root:/${encodePath(tryPath)}?$select=id,name,folder,webUrl`)
                .get();
              if (folderItem?.id && folderItem.folder) {
                folder = folderItem;
                console.log(`[VesselDMS] _uploadFileToFolder: resolved folder in SPO at "${tryPath}" (id=${folder.id})`);
                break;
              }
            } catch {
              // Check next candidate
            }
          }

          // 2. If folder does not exist yet in SPO, recursively create every missing segment
          if (!folder) {
            const targetPathToCreate = targetSiteId && cleanDirectPath ? cleanDirectPath : canonicalSharePointPath;
            console.log(`[VesselDMS] _uploadFileToFolder: ensuring folder path "${targetPathToCreate}" in SharePoint Online...`);
            const segments = targetPathToCreate.split('/').filter(Boolean);
            let currentParentPath = '';
            let currentFolderItem: any = null;

            for (let i = 0; i < segments.length; i++) {
              const seg = segments[i].trim();
              const accumulatedPath = currentParentPath ? `${currentParentPath}/${seg}` : seg;
              const encodedAccumulated = encodePath(accumulatedPath);

              try {
                const existing = await graphClient
                  .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/root:/${encodedAccumulated}?$select=id,name,folder,webUrl`)
                  .get();
                if (existing?.id && existing.folder) {
                  currentFolderItem = existing;
                  currentParentPath = accumulatedPath;
                  continue;
                }
              } catch {
                // Not found, create it
              }

              // Never auto-create a department or Kaizen root folder
              // (Technical & Crewing, Commercial & Chartering, Insurance,
              // Kaizen - Knowledge Bank). If it isn't already in SharePoint,
              // fail the upload instead of building the template tree.
              if (!currentParentPath && (
                this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === seg.toLowerCase()) ||
                /^kaizen - knowledge bank$/i.test(seg)
              )) {
                throw new Error(`Folder "${seg}" does not exist in SharePoint Online and will not be created automatically.`);
              }

              const createUrl = currentParentPath
                ? `/sites/${effectiveSiteId}/drives/${effectiveDriveId}/root:/${encodePath(currentParentPath)}:/children`
                : `/sites/${effectiveSiteId}/drives/${effectiveDriveId}/root/children`;

              try {
                const created = await graphClient.api(createUrl).post({
                  name: seg,
                  folder: {},
                  '@microsoft.graph.conflictBehavior': 'fail',
                });
                currentFolderItem = created;
                currentParentPath = accumulatedPath;
                console.log(`[VesselDMS] _uploadFileToFolder: created SPO folder "${seg}" (id=${created.id})`);
              } catch (createErr: any) {
                try {
                  const existingAfterConflict = await graphClient
                    .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/root:/${encodedAccumulated}?$select=id,name,folder,webUrl`)
                    .get();
                  currentFolderItem = existingAfterConflict;
                  currentParentPath = accumulatedPath;
                } catch {
                  throw createErr;
                }
              }
            }
            folder = currentFolderItem;
          }
        }

        if (!folder?.id) {
          throw new Error(`Could not locate or create SharePoint folder for "${subFolderPath || uploadFolderId}".`);
        }

        // 2.5 Block same-name duplicates within this folder/site. A vessel's
        // folder lives on exactly one SharePoint site's drive, so this check
        // is naturally scoped per site — the same file name is still allowed
        // in a different site (or a different vessel folder on the same
        // site) because that's a different `folder.id` entirely and this
        // lookup never sees it. Previously nothing blocked this: a small
        // file (<=4MB) silently overwrote the existing one via the default
        // "replace" conflict behavior on the simple upload endpoint, and a
        // large file (>4MB) got silently auto-renamed ("file (1).pdf") by
        // the chunked upload session's conflictBehavior: 'rename' — neither
        // path ever told the user a duplicate existed.
        try {
          const existingFile = await graphClient
            .api(`/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}:/${encodeURIComponent(file.name)}?$select=id,name,file,webUrl`)
            .get();
          if (existingFile?.id && existingFile.file) {
            console.warn('[VesselDMS] Duplicate-name guard: existing file found', {
              name: existingFile.name, webUrl: existingFile.webUrl, folderId: folder.id, folderName: folder.name,
              siteId: effectiveSiteId, driveId: effectiveDriveId,
            });
            throw new Error(
              `A file named "${file.name}" already exists in this folder${existingFile.webUrl ? ` (${existingFile.webUrl})` : ''}. Rename the file, delete the existing one, or upload it to a different vessel/folder before trying again.`
            );
          }
        } catch (dupCheckErr: any) {
          // A thrown duplicate error above must propagate; a 404 ("Not
          // found") from the existence check itself just means no
          // conflict — let the upload proceed.
          if (dupCheckErr instanceof Error && dupCheckErr.message.startsWith('A file named')) {
            throw dupCheckErr;
          }
        }

        // 3. Upload file directly into the resolved SharePoint Online folder
        let item: any = null;
        if (file.size <= 4 * 1024 * 1024) {
          const uploadUrl = `/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}:/${encodeURIComponent(file.name)}:/content?@microsoft.graph.conflictBehavior=fail`;
          item = await graphClient.api(uploadUrl).put(file);
        } else {
          // Large file chunked upload session (> 4 MB). conflictBehavior is
          // 'fail' (not the previous 'rename') so a race with another
          // upload of the same name is rejected by SharePoint itself rather
          // than silently creating "file (1).pdf".
          const sessionUrl = `/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}:/${encodeURIComponent(file.name)}:/createUploadSession`;
          const session = await graphClient.api(sessionUrl).post({
            item: { '@microsoft.graph.conflictBehavior': 'fail', name: file.name },
          });
          const uploadSessionUrl = session.uploadUrl;
          const CHUNK_SIZE = 320 * 1024 * 10; // 3.2 MB chunks
          let offset = 0;
          while (offset < file.size) {
            const end = Math.min(offset + CHUNK_SIZE, file.size);
            const chunk = file.slice(offset, end);
            const res = await fetch(uploadSessionUrl, {
              method: 'PUT',
              headers: {
                'Content-Length': String(end - offset),
                'Content-Range': `bytes ${offset}-${end - 1}/${file.size}`,
              },
              body: chunk,
            });
            if (res.status === 200 || res.status === 201) {
              item = await res.json();
              break;
            } else if (res.status !== 202) {
              const errTxt = await res.text().catch(() => res.statusText);
              throw new Error(`Upload chunk error (${res.status}): ${errTxt}`);
            }
            offset = end;
          }
        }

        if (!item?.id) throw new Error('SharePoint did not return an uploaded file ID.');
        console.log(`[VesselDMS] _uploadFileToFolder: successfully uploaded "${file.name}" to SharePoint Online (id=${item.id}) in folder "${folder.name}"`);
        this._appUploadedFileIds.add(item.id as string);

        // Optimistically update site folder cache & live items
        const siteCacheKey = `${effectiveSiteId}::${effectiveDriveId}::${folder.id}`;
        const newFileItem = {
          id: item.id,
          name: file.name,
          file: {},
          size: file.size,
          lastModifiedDateTime: new Date().toISOString(),
          webUrl: item.webUrl || '',
        };
        const existingSiteCache = this._siteFolderItemsCache.get(siteCacheKey);
        if (existingSiteCache) {
          if (!existingSiteCache.items.some(i => i.name.toLowerCase() === file.name.toLowerCase())) {
            existingSiteCache.items = [...existingSiteCache.items, newFileItem];
          }
        } else {
          this._siteFolderItemsCache.set(siteCacheKey, { items: [newFileItem], loading: false, parentPath: '' });
        }
        void this._refreshSiteFolder(effectiveSiteId, effectiveDriveId, folder.id);

        return { fileId: item.id as string, statusPending: false, folderId: folder.id as string, isGraphUpload: true, webUrl: item.webUrl || null };
      } catch (graphErr) {
        const detail = graphErr instanceof Error ? graphErr.message : String(graphErr);
        console.error('[VesselDMS] Graph direct upload error:', detail);
        throw new Error(`Could not upload file to SharePoint Online: ${detail}`);
      }
    }

    // REST backend fallback (no Graph client configured).
    const form = new FormData();
    form.append('file', file);
    if (this.props.userEmail) form.append('uploader_email', this.props.userEmail);
    if (this.state.activeDocumentSite) form.append('site_key', this.state.activeDocumentSite);

    const liveFolderId = this._getLiveSharePointFolderId(canonicalSharePointPath) || this._getLiveSharePointFolderId(subFolderPath);
    const restFolderId = liveFolderId || (uploadFolderId && !uploadFolderId.startsWith('sf_') ? uploadFolderId : canonicalSharePointPath);

    let endpoint: string;
    if (restFolderId.includes('/')) {
      endpoint = `${base}/api/folders/upload-by-path?path=${encodeURIComponent(restFolderId)}`;
    } else {
      endpoint = monthDriven
        ? `${base}/api/folders/${encodeURIComponent(restFolderId)}/month-upload`
        : `${base}/api/folders/${encodeURIComponent(restFolderId)}/upload`;
    }

    const resp = await fetch(endpoint, {
      method: 'POST',
      headers: this._uploadHeaders(),
      body: form,
    });

    if (!resp.ok) {
      const errText = await resp.text().catch(() => resp.statusText);
      throw new Error(`HTTP ${resp.status}: ${errText}`);
    }

    const data = await resp.json().catch(() => ({}));
    if (data?.drive_item_id) this._appUploadedFileIds.add(String(data.drive_item_id));
    // REST upload: folderId is a backend DB ID (NOT a SharePoint drive item ID),
    // so isGraphUpload must be false — a downstream refresh must resolve the real
    // SPO folder ID from the path rather than querying Graph with this DB ID.
    return {
      fileId: data?.id || null,
      statusPending: data?.status === 'pending',
      folderId: restFolderId || null,
      isGraphUpload: false,
      webUrl: data?.webUrl || data?.web_url || null,
    };
  }

  /**
   * Fire-and-forget OCR staging trigger:
   * Queues an uploaded file into the unified OCR staging pipeline for classification,
   * tag suggestion from the Templates module, and user review before moving.
   * Called for both folder uploads (source='folder') and single-file uploads (source='direct').
   */
  public _triggerOcrStaging(
    driveItemId: string | null,
    file: File,
    folderId: string,
    subFolderPath: string,
    vesselName: string,
    source: 'folder' | 'direct' = 'direct'
  ): void {
    // Disabled per user request: we do not want review & ocr
    return;
  }

  // ── Graph API Folder Walking (mirrors VesselListView.tsx from reference project) ────

  public readonly GRAPH_FETCH_CONCURRENCY = 6;
  // Note: VESSEL_ROOT and MAIN_FOLDER_NAMES are defined as class fields above.
  // Primary structure: Documents/{MainFolder}/{VesselName}/...

  /**
   * List children of a SharePoint drive path via Graph API.
   * Returns an array of { id, name, isFolder, monthDriven, upload } items.
   */
  public async _getGraphChildren(
    folderPath: string,
    signal: AbortSignal,
  ): Promise<Array<{ id: string; name: string; size?: number; isFolder: boolean; monthDriven: boolean; upload: boolean; createdDateTime?: string; lastModifiedDateTime?: string }>> {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return [];

    const fetchChildrenForPath = async (p: string) => {
      const cleanP = p.replace(/^\/+/, '');
      const encodedPath = cleanP.split('/').map(s => encodeURIComponent(s)).join('/');
      // Graph root children uses /root/children, not /root:/:/children.
      const itemPath = cleanP ? `root:/${encodedPath}:/children` : 'root/children';
      const url = `/sites/${siteId}/drives/${driveId}/${itemPath}` +
        `?$select=id,name,size,folder,file,createdDateTime,lastModifiedDateTime&$top=200`;
      const result: any = await this._getGraphChildrenResponse(url);
      if (signal.aborted) return [];
      return (result.value ?? []).map((item: any) => ({
        id: item.id as string,
        name: item.name as string,
        size: typeof item.size === 'number' ? item.size : undefined,
        isFolder: !!item.folder,
        monthDriven: false,
        upload: !!item.folder,
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime,
      }));
    };

    try {
      return await fetchChildrenForPath(folderPath);
    } catch (err: any) {
      const status = err?.statusCode ?? err?.code ?? 0;
      if (status === 404) return [];
      console.warn('[VesselDMS] _getGraphChildren failed for path:', folderPath, err);
      return [];
    }
  }


  public _mapToStandardGroup(rawName: string): string | null {
    const n = (rawName || '').replace(/^Folder-\d+\s+/i, '').trim().toLowerCase();
    if (n === 'technical & crewing' || n === 'technical' || n.includes('technical and crewing') || n.includes('technical & crewing')) {
      return 'Technical & Crewing';
    }
    if (n === 'commercial & chartering' || n.includes('commercial and chartering') || n.includes('commercial & chartering') || n === 'commercial') {
      return 'Commercial & Chartering';
    }
    if (n === 'insurance' || n.includes('insurance')) {
      return 'Insurance';
    }
    if (n.includes('kaizen')) {
      return 'Kaizen - Knowledge Bank';
    }
    return null;
  }

  /**
   * Recursively walk a Graph folder path and emit FlatRow entries.
   * Mirrors the walkFolder() + leafToRows() pattern from the reference project's VesselListView.tsx.
   */
  public async _walkGraphFolder(
    folderPath: string,
    pathParts: string[],
    vesselName: string,
    signal: AbortSignal,
    srCounter: { value: number },
    onRows: (rows: FlatRow[]) => void,
    folderId?: string,
  ): Promise<void> {
    if (signal.aborted) return;

    let kids: Array<{ id: string; name: string; size?: number; isFolder: boolean; monthDriven: boolean; upload: boolean; createdDateTime?: string; lastModifiedDateTime?: string }>;
    try {
      kids = await this._getGraphChildren(folderPath, signal);
    } catch {
      return;
    }
    if (signal.aborted) return;

    const files = kids.filter(k => !k.isFolder);
    const subFolders = kids.filter(k => k.isFolder);
    // Strip "Folder-X " prefix (e.g. "Folder-1 Technical & Crewing" → "Technical & Crewing")
    const stripPrefix = (s: string): string => s.replace(/^Folder-\d+\s+/i, '');

    const isKaizen = vesselName.toLowerCase().includes('kaizen');
    const mappedGroup = this._mapToStandardGroup(pathParts[0] || '');
    let group = mappedGroup || stripPrefix(pathParts[0] || '');
    let category = pathParts.length >= 3 ? stripPrefix(pathParts[2]) : (pathParts.length >= 2 ? stripPrefix(pathParts[pathParts.length - 1]) : group);
    let subCategory = pathParts.length >= 4 ? stripPrefix(pathParts[pathParts.length - 1]) : category;
    let subPath = '';

    if (isKaizen) {
      group = 'Kaizen - Knowledge Bank';
      category = pathParts.length >= 1 ? stripPrefix(pathParts[0]) : group;
      subCategory = pathParts.length >= 2 ? stripPrefix(pathParts[pathParts.length - 1]) : category;
      subPath = ['Kaizen - Knowledge Bank', ...pathParts.map(stripPrefix)].join(' > ');
    } else {
      // pathParts layout: [group, vesselName, category, subCategory1, subCategory2, ...]
      const categories = pathParts.length >= 2 ? pathParts.slice(2).map(stripPrefix) : pathParts.map(stripPrefix);
      const breadcrumbParts = [vesselName, group, ...categories];
      subPath = breadcrumbParts.join(' > ');
    }

    const groupKey = `${vesselName}||${group}||${category}||${subCategory}||${subPath}`;
    const canUpload = subFolders.length === 0; // leaf nodes can receive uploads

    // Emit a leaf row (matches reference leafToRows pattern)
    if (files.length > 0 || subFolders.length === 0) {
      srCounter.value += 1;
      const baseSr = String(srCounter.value);
      const suffixes = 'abcdefghijklmnopqrstuvwxyz';

      if (files.length === 0) {
        onRows([{
          srNo: baseSr, vesselName, group, category, subCategory,
          subFolderPath: subPath, fileName: null, fileId: null,
          canUpload, groupKey,
          uploadFolderId: folderId || folderPath,
          monthDriven: false,
        }]);
      } else {
        const leafRows: FlatRow[] = files.map((f, idx) => ({
          srNo: idx === 0 ? baseSr : `${baseSr}${suffixes[idx - 1] ?? idx}`,
          vesselName, group, category, subCategory, subFolderPath: subPath,
          fileName: f.name, fileId: f.id, canUpload, groupKey,
          uploadFolderId: folderId || folderPath,
          fileUploadedAt: f.createdDateTime ? Date.parse(f.createdDateTime) : (f.lastModifiedDateTime ? Date.parse(f.lastModifiedDateTime) : undefined),
          fileSize: typeof (f as any).size === 'number' ? `${((f as any).size / 1024).toFixed(1)} KB` : undefined,
          monthDriven: false,
        }));
        onRows(leafRows);
      }
    }

    // Recurse into subfolders with bounded concurrency
    // pathParts layout: [mainFolder, vesselName, category, subCategory, ...]
    if (subFolders.length > 0 && !signal.aborted) {
      await this._mapLimit(
        subFolders,
        this.GRAPH_FETCH_CONCURRENCY,
        async sf => this._walkGraphFolder(
          `${folderPath}/${sf.name}`,
          [...pathParts, sf.name],
          vesselName,
          signal,
          srCounter,
          onRows,
          sf.id,
        ),
      );
    }
  }

  /** Bounded-concurrency async map (mirrors mapLimit from reference project). */
  public async _mapLimit<T, R>(
    items: T[],
    limit: number,
    mapper: (item: T, index: number) => Promise<R>,
  ): Promise<R[]> {
    if (items.length === 0) return [];
    const results: R[] = new Array(items.length);
    let nextIndex = 0;
    const runWorker = async (): Promise<void> => {
      let i = nextIndex++;
      while (i < items.length) {
        results[i] = await mapper(items[i], i);
        i = nextIndex++;
      }
    };
    const workers = Array.from({ length: Math.min(limit, items.length) }, runWorker);
    await Promise.all(workers);
    return results;
  }

  /**
   * Flatten a single vessel's SPO folder tree via Graph API.
   * Calls onChunk with each batch of rows as they arrive (progressive rendering).
   */
  public async _flattenVesselViaGraph(
    vesselName: string,
    signal: AbortSignal,
    onChunk: (rows: FlatRow[]) => void,
  ): Promise<FlatRow[]> {
    const allRows: FlatRow[] = [];
    const srCounter = { value: 0 };

    // Discover the vessel's actual location before walking it. This avoids
    // issuing Graph requests for every possible template branch and supports
    // both the current Vessels/{vessel} structure and legacy main-folder
    // structures already present in SharePoint.
    const normaliseName = (name: string): string => cleanName(name || '')
      .replace(/^Folder-\d+\s+/i, '')
      .trim()
      .toLowerCase();
    const knownMainFolders = new Set(this.MAIN_FOLDER_NAMES.map(normaliseName));
    const discoveredMainFolders: Array<{ path: string; group: string }> = [];
    let foundSupportedStructure = false;
    const basePaths = [''];

    // Primary structure: {MainFolder} / {Vessel} directly at root
    const isKaizen = normaliseName(vesselName).includes('kaizen');
    const rootChildren = await this._getGraphChildren('', signal).catch(() => []);

    if (isKaizen) {
      foundSupportedStructure = true;
      const kaizenNode = rootChildren.find(node => node.isFolder && normaliseName(node.name).includes('kaizen'));
      const kaizenPath = kaizenNode ? kaizenNode.name : 'Kaizen - Knowledge Bank';
      await this._walkGraphFolder(kaizenPath, [], vesselName, signal, srCounter, rows => {
        allRows.push(...rows);
        onChunk(rows);
      });
      return allRows;
    }

    // Check all known main folders first (including flexible naming like "Technical and Crewing  New")
    const mainFolderNodes = rootChildren.filter(node => node.isFolder && (knownMainFolders.has(normaliseName(node.name)) || this._mapToStandardGroup(node.name) !== null));
    if (mainFolderNodes.length > 0) {
      foundSupportedStructure = true;
      for (const mainNode of mainFolderNodes) {
        if (signal.aborted) break;
        const mainPath = mainNode.name;
        const standardGroup = this._mapToStandardGroup(mainNode.name) || mainNode.name;
        const mainChildren = await this._getGraphChildren(mainPath, signal).catch(() => []);
        const vesselNode = mainChildren.find(node => node.isFolder && normaliseName(node.name) === normaliseName(vesselName));
        if (vesselNode) {
          discoveredMainFolders.push({ path: `${mainPath}/${vesselNode.name}`, group: standardGroup });
        }
      }
    }

    // Legacy fallback: Vessels / Specific Vessels / {Vessel} / {MainFolder} or Vessels / {Vessel} / {MainFolder}
    if (discoveredMainFolders.length === 0) {
    for (const basePath of basePaths) {
      if (signal.aborted || discoveredMainFolders.length > 0) break;
      const baseChildren = await this._getGraphChildren(basePath, signal).catch(() => []);
      const prefix = basePath ? `${basePath}/` : '';

      // Current structure: Vessels / Specific Vessels / {Vessel} / {Main folder} or Vessels / {Vessel} / {Main folder}.
      const vesselsRoot = baseChildren.find(node => node.isFolder && normaliseName(node.name) === 'vessels');
      if (vesselsRoot) {
        foundSupportedStructure = true;
        const vesselsChildren = await this._getGraphChildren(`${prefix}${vesselsRoot.name}`, signal).catch(() => []);

        // 1. Check under Specific Vessels/
        const specificVesselsNode = vesselsChildren.find(node => node.isFolder && normaliseName(node.name) === 'specific vessels');
        let vesselNode = null;
        let vesselParentPath = '';
        if (specificVesselsNode) {
          const specificVesselNodes = await this._getGraphChildren(`${prefix}${vesselsRoot.name}/${specificVesselsNode.name}`, signal).catch(() => []);
          vesselNode = specificVesselNodes.find(node => node.isFolder && normaliseName(node.name) === normaliseName(vesselName));
          if (vesselNode) {
            vesselParentPath = `${prefix}${vesselsRoot.name}/${specificVesselsNode.name}`;
          }
        }

        // 2. Check directly under Vessels/
        if (!vesselNode) {
          vesselNode = vesselsChildren.find(node => node.isFolder && normaliseName(node.name) === normaliseName(vesselName));
          if (vesselNode) {
            vesselParentPath = `${prefix}${vesselsRoot.name}`;
          }
        }

        if (vesselNode && vesselParentPath) {
          const vesselPath = `${vesselParentPath}/${vesselNode.name}`;
          const mainNodes = await this._getGraphChildren(vesselPath, signal).catch(() => []);
          mainNodes.filter(node => node.isFolder).forEach(node => {
            const standardGroup = this._mapToStandardGroup(node.name) || node.name;
            discoveredMainFolders.push({ path: `${vesselPath}/${node.name}`, group: standardGroup });
          });
        }
      }

      // Legacy structure: {Main folder} / {Vessel} / ... or
      // {Main folder} / Vessels / {Vessel} / ....
      if (discoveredMainFolders.length === 0) {
        const mainNodes = baseChildren.filter(node => node.isFolder && (knownMainFolders.has(normaliseName(node.name)) || this._mapToStandardGroup(node.name) !== null));
        if (mainNodes.length > 0) foundSupportedStructure = true;
        for (const mainNode of mainNodes) {
          if (signal.aborted) break;
          const mainPath = `${prefix}${mainNode.name}`;
          const standardGroup = this._mapToStandardGroup(mainNode.name) || mainNode.name;
          const mainChildren = await this._getGraphChildren(mainPath, signal).catch(() => []);
          let vesselNode = mainChildren.find(node => node.isFolder && normaliseName(node.name) === normaliseName(vesselName));
          let vesselPath = vesselNode ? `${mainPath}/${vesselNode.name}` : '';
          if (!vesselNode) {
            const legacyVesselsRoot = mainChildren.find(node => node.isFolder && normaliseName(node.name) === 'vessels');
            if (legacyVesselsRoot) {
              const vesselNodes = await this._getGraphChildren(`${mainPath}/${legacyVesselsRoot.name}`, signal).catch(() => []);
              vesselNode = vesselNodes.find(node => node.isFolder && normaliseName(node.name) === normaliseName(vesselName));
              vesselPath = vesselNode ? `${mainPath}/${legacyVesselsRoot.name}/${vesselNode.name}` : '';
            }
          }
          if (vesselNode && vesselPath) discoveredMainFolders.push({ path: vesselPath, group: standardGroup });
        }
      }
    }
    } // end legacy fallback

    if (discoveredMainFolders.length > 0) {
      await this._mapLimit(discoveredMainFolders, 3, async folder => {
        await this._walkGraphFolder(folder.path, [folder.group, vesselName], vesselName, signal, srCounter, rows => {
          allRows.push(...rows);
          onChunk(rows);
        });
      });
      return allRows;
    }

    // A recognised vessel layout exists, but this vessel does not exist in
    // it. Do not probe every template path: those requests only create noisy
    // Graph 404s and cannot produce files for the requested vessel.
    if (foundSupportedStructure) return allRows;

    // Root-level names already confirmed to exist (from the single
    // rootChildren fetch above, via mainFolderNodes at the top of this
    // function). A mainFolderName NOT in this set is proven absent from the
    // site's root, so `{mainFolderName}/...` legacy guesses below can only
    // ever 404 — skip them instead of spending 2 guaranteed-404 Graph calls
    // per main folder per vessel on a site that doesn't use that layout at
    // all (this is what was flooding sites with a genuinely different
    // top-level structure, e.g. no Technical & Crewing/Commercial &
    // Chartering/Insurance folders at root).
    const confirmedRootMainFolders = new Set(mainFolderNodes.map(node => normaliseName(node.name)));

    // Walk each main folder in parallel
    await this._mapLimit(
      this.MAIN_FOLDER_NAMES,
      3,
      async mainFolderName => {
        if (signal.aborted) return;
        // Real layout is always {main}/{vessel} (docs/folder-structure-mode.md
        // "Layout"; matches every other path built in this file, e.g. L5690,
        // L5733, L5745). We reach this branch only when `foundSupportedStructure`
        // is false, which means `confirmedRootMainFolders` is already known to
        // be empty (it's derived from the same root scan) — so *no* probe for
        // this mainFolderName can ever succeed. Skip Graph entirely instead of
        // firing guaranteed-404 requests; the template fallback below (topCats
        // empty) renders the placeholder rows either way.
        let vesselPath = `${this.VESSEL_ROOT}/${mainFolderName}/${vesselName}`;
        let topCats: Awaited<ReturnType<typeof this._getGraphChildren>> = [];
        if (confirmedRootMainFolders.has(normaliseName(mainFolderName))) {
          const legacyPaths = [
            `${mainFolderName}/${vesselName}`,
            `${mainFolderName}/Vessels/${vesselName}`,
          ];
          for (const legacyPath of legacyPaths) {
            const legacyCats = await this._getGraphChildren(legacyPath, signal).catch(() => []);
            if (legacyCats.length > 0) {
              vesselPath = legacyPath;
              topCats = legacyCats;
              break;
            }
          }
        }
        if (signal.aborted) return;

        const folders = topCats.filter(c => c.isFolder);
        const files = topCats.filter(c => !c.isFolder);

        if (folders.length === 0) {
          // Vessel subfolders not yet provisioned in SPO — expand standard DMS template hierarchy
          const displayMainFolder = mainFolderName.replace(/^Folder-\d+\s+/i, '');
          const templateTree: Record<string, Array<{ category: string; subCats: string[] }>> = Object.fromEntries(
            MAIN_FOLDERS.map(main => [main.name, main.perVesselTree.map(folder => {
              const leaves: string[] = [];
              const collectLeaves = (node: { name: string; children?: Array<{ name: string; children?: any[] }> }, path: string[]): void => {
                const nextPath = [...path, node.name];
                if (!node.children?.length) leaves.push(nextPath.join(' > '));
                else node.children.forEach(child => collectLeaves(child, nextPath));
              };
              collectLeaves(folder, []);
              return { category: folder.name, subCats: leaves };
            })])
          );

          const cats = templateTree[displayMainFolder] || [
            { category: displayMainFolder, subCats: [displayMainFolder] }
          ];

          const fallbackRows: FlatRow[] = [];
          for (const item of cats) {
            for (const subCat of item.subCats) {
              srCounter.value += 1;
              const nestedPath = subCat.split(' > ');
              const breadcrumb = `${vesselName} > ${displayMainFolder} > ${nestedPath.join(' > ')}`;
              const logicalPath = `${vesselName}/${displayMainFolder}/${nestedPath.join('/')}`;
              fallbackRows.push({
                srNo: String(srCounter.value),
                vesselName,
                group: displayMainFolder,
                category: item.category,
                subCategory: subCat,
                subFolderPath: breadcrumb,
                fileName: null,
                fileId: null,
                canUpload: true,
                groupKey: `${vesselName}||${displayMainFolder}||${item.category}||${subCat}||${breadcrumb}`,
                uploadFolderId: logicalPath,
                monthDriven: false,
              });
            }
          }

          allRows.push(...fallbackRows);
          onChunk(fallbackRows);
          return;
        }

        // Walk each category folder with bounded concurrency
        await this._mapLimit(
          folders,
          this.GRAPH_FETCH_CONCURRENCY,
          async cat => {
            if (signal.aborted) return;
            const chunkRows: FlatRow[] = [];
            await this._walkGraphFolder(
              `${vesselPath}/${cat.name}`,
              [mainFolderName, vesselName, cat.name],
              vesselName,
              signal,
              srCounter,
              rows => {
                chunkRows.push(...rows);
                allRows.push(...rows);
              },
            );
            if (chunkRows.length > 0) onChunk(chunkRows);
          },
        );
      },
    );

    return allRows;
  }

  /**
   * Flatten ALL vessels via Graph API (called from _loadData when Graph context is available).
   * Streams rows progressively into state via onChunk for a responsive UI.
   */
  public async _flattenAllViaGraph(vessels: VesselRecord[], signal: AbortSignal): Promise<FlatRow[]> {
    const allRows: FlatRow[] = [];

    await this._mapLimit(
      vessels,
      3,
      async vessel => {
        if (signal.aborted) return;
        await this._flattenVesselViaGraph(
          vessel.name,
          signal,
          chunk => {
            allRows.push(...chunk);
          },
        );
      },
    );

    return allRows;
  }

  public async _flattenAll(vessels: VesselRecord[], signal: AbortSignal): Promise<FlatRow[]> {
    const base = this._base(); const out: FlatRow[] = []; let sr = 0;
    for (const v of vessels) {
      if (signal.aborted) break;
      try {
        const mains = await this._fetchJson(`${base}/api/vessels/${v.id}/mains`, signal);
        for (const m of mains) {
          if (signal.aborted) break;
          const group = m.name;
          try {
            const cats = await this._fetchJson(`${base}/api/folders/${m.id}/children`, signal);
            for (const cat of cats) {
              if (cat?.kind === 'file') continue;
              // category = folder directly under vessel
              const category = cat.name;
              // Try to get sub-folders (sub-categories)
              try {
                const subCats = await this._fetchJson(`${base}/api/folders/${cat.id}/children`, signal);
                const leaves = subCats.filter((s: any) => s?.kind !== 'file');
                if (leaves.length > 0) {
                  for (const leaf of leaves) {
                    if (leaf?.kind === 'file') continue; sr++;
                    const subCategory = leaf.name;
                    const sp = `${v.name} > ${group} > ${category} > ${subCategory}`;
                    out.push({ srNo: String(sr), vesselName: v.name, group, category, subCategory, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: `${v.id}||${group}||${category}||${subCategory}||${sp}`, uploadFolderId: leaf.id, monthDriven: Boolean(leaf.month_driven) });
                  }
                } else {
                  // No sub-folders → cat itself is a leaf
                  sr++;
                  const sp = `${v.name} > ${group} > ${category}`;
                  out.push({ srNo: String(sr), vesselName: v.name, group, category, subCategory: category, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: `${v.id}||${group}||${category}||${category}||${sp}`, uploadFolderId: cat.id, monthDriven: Boolean(cat.month_driven) });
                }
              } catch {
                sr++;
                const sp = `${v.name} > ${group} > ${category}`;
                out.push({ srNo: String(sr), vesselName: v.name, group, category, subCategory: category, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: `${v.id}||${group}||${category}||${category}||${sp}`, uploadFolderId: cat.id, monthDriven: Boolean(cat.month_driven) });
              }
            }
          } catch { /* skip */ }
        }
      } catch { /* skip */ }
    }
    return out;
  }

  // ── Navigation & Views ────────────────────────────────────────────────────

  /** Push a new folder navigation entry, truncating any forward history. */
  public _pushFolderNav(
    folderPathStack: { id: string; name: string }[],
    docMainFolder: State['docMainFolder'],
  ): void {
    this.setState(prev => {
      const truncated = prev.folderNavHistory.slice(0, prev.folderNavIndex + 1);
      const next = [...truncated, { folderPathStack, docMainFolder }];
      return { folderNavHistory: next, folderNavIndex: next.length - 1, folderPathStack, docMainFolder };
    });
  }

  public _goToView = async (view: AppView): Promise<void> => {
    // Settings → Module Management can hide a module app-wide; refuse to
    // navigate into one even if the caller (a stale bookmark, a dashboard
    // shortcut, direct state) still asks for it. Settings/Profile are never
    // in hiddenModules (module_settings_api excludes them), so this never
    // blocks getting back to Settings to re-enable something.
    if (this.state.hiddenModules.indexOf(view) !== -1) {
      view = 'dashboard';
    }
    const previousView = this.state.view;
    // Picking any app section (sidebar, deep links) leaves the embedded
    // traditional SharePoint site.
    this.setState({ view, classicSiteOpen: false });
    if (view === 'dashboard') {
      void this._loadDashboardStats();
    } else if (view === 'list') {
      // 'list' is the Documents module (Sidebar.tsx navItems: { id: 'list',
      // label: 'Documents' }). This branch only ever runs from that one
      // sidebar button (renderNavBtn's onClick -> _goToView(item.id)) — no
      // other call site in the app passes 'list', and none of Documents'
      // OWN internal navigation (breadcrumbs, clicking into a subfolder,
      // Folder<->List view toggle, "load more vessels") goes through
      // _goToView at all, so this can't fire mid-browse. That's what makes
      // it safe to treat as "fresh entry into the module" and reset the
      // filter bar: a vessel/group/category/attachment/search filter left
      // over from a previous visit was showing a stale filtered view (and
      // a vessel dropdown pinned to a vessel the user wasn't even looking
      // at) the next time Documents was opened, because this component is
      // a single persistent instance for the whole web part — switching
      // sidebar sections never unmounts/remounts it, so nothing was ever
      // clearing these fields on its own.
      //
      // activeDocumentSite is left alone here — the "SharePoint site"
      // dropdown selection isn't part of the root breadcrumb and switching
      // it has its own dedicated flow (_switchDocumentSite) that resets its
      // own filters. folderPathStack / docMainFolder / docScopeType /
      // docViewMode ARE reset below (see the note further down): a click on
      // this sidebar button is meant to be a clean re-entry into the module.
      //
      // No separate refetch is needed to show the unfiltered contents:
      // `rows` already holds the unfiltered data (filtering happens
      // entirely client-side, computed fresh on every render — see
      // DocumentsPage.tsx's `filtered`), so clearing these fields is
      // sufficient for the next render to show the current folder's real
      // contents instead of the stale filtered subset.
      //
      // folderPathStack / docMainFolder / docScopeType / docViewMode ARE
      // reset here too (this used to be the one exception, see the removed
      // comment above this block in an earlier revision): the sidebar
      // "Documents" button is meant to be a clean re-entry point into the
      // module, not a "resume where I left off" shortcut. The sessionStorage
      // restore in the constructor (_readPersistedFolderNav) still runs on
      // an actual page refresh, so `view` starts back at 'dashboard' with
      // the last folder pre-loaded in state — but nothing ever surfaces that
      // restored folder until this branch runs (view only ever becomes
      // 'list' via this function), so leaving these fields untouched here
      // meant a page refresh + a single click into Documents always dropped
      // the user back into whatever deep folder they'd last browsed instead
      // of the module's root. Resetting them here makes the click
      // deterministic; componentDidUpdate's persistence effect clears the
      // now-stale sessionStorage entry as soon as folderPathStack changes.
      //
      // EXCEPTION: when the user was actively browsing a live SharePoint
      // site/library ('sites' / 'shared_docs' / 'documents' — real Graph
      // folder data behind activeDocumentSite, not the legacy vessels DB
      // tree), stomping folderPathStack/docScopeType back to the vessels
      // root threw that browsing session away on every single re-entry —
      // after a page refresh (sessionStorage restores the live scope into
      // state, then the first "Documents" click here wiped it straight back
      // out) and after switching to any other sidebar module and back. The
      // user saw this as "files never load" / "only folders show up": they
      // were dropped into the generic vessels-DB root (department tiles,
      // no live files) instead of the real site/folder they were just
      // looking at, and had to re-pick the site from the dropdown every
      // time. A live scope's own navigation (_switchDocumentSite, folder
      // clicks) already resets its own filters/breadcrumb when the user
      // deliberately changes site or drills in, so it doesn't need this
      // reset to avoid a "stale filtered view" — only the legacy
      // vessels/kaizen/common scopes do.
      //
      // FURTHER EXCEPTION TO THE EXCEPTION: the live-scope state above can
      // also be true on the very *first* "Documents" click after a hard
      // refresh, purely because the constructor seeded it straight from
      // sessionStorage (_readPersistedFolderNav) — the user never actually
      // browsed anywhere this page load, they just had a deep folder/site
      // left over from before the refresh. Treating that as "resume my live
      // browsing session" drops them right back into a folder that may no
      // longer have anything freshly loaded for it (empty-looking view) and
      // defeats the point of a hard refresh. _hasEnteredDocumentsSinceMount
      // distinguishes the two cases: false only until the user's own first
      // click into Documents since this component was constructed, so a
      // page-load-only restore never counts as "was live browsing," while
      // switching sidebar tabs and coming back mid-session still does.
      const isFirstEntrySinceMount = !this._hasEnteredDocumentsSinceMount;
      this._hasEnteredDocumentsSinceMount = true;
      const isLiveScope =
        this.state.docScopeType === 'sites' ||
        this.state.docScopeType === 'shared_docs' || this.state.docScopeType === 'documents';
      // On the very first click, only trust a live scope that
      // _loadDocumentSites has actually finished verifying against live
      // Graph data (_initialDocumentsRootReady) — otherwise it really could
      // be an unverified sessionStorage restore, and the reset below is
      // still the right call. If _loadDocumentSites already finished, this
      // is either a freshly live-verified root or a re-synced restore
      // (see its comments), so treat it the same as later re-entries
      // instead of throwing it away and landing on the empty vessels root.
      const wasLiveLibraryScope = isLiveScope && (!isFirstEntrySinceMount || this._initialDocumentsRootReady);
      if (wasLiveLibraryScope) {
        // First entry since a page load: the breadcrumb was seeded from
        // sessionStorage (a vessel folder browsed before the refresh) while the
        // filters below are reset to "All vessels" / "All main folders", so the
        // two disagreed (breadcrumb showed an old vessel with All vessels
        // selected). Land on the library root (site + drive) instead; later
        // re-entries in the same session keep the live browsing position.
        const stack = this.state.folderPathStack || [];
        const rootLen = stack[0]?.id === 'sites_root' ? 3
          : ((stack[0]?.id === 'lib:shared_documents' || stack[0]?.id === 'lib:documents') ? 1 : -1);
        const resetToRoot = isFirstEntrySinceMount && rootLen >= 0 && stack.length > rootLen;
        const rootStack = resetToRoot ? stack.slice(0, rootLen) : stack;
        this.setState({
          textFilter: '',
          vesselFilter: 'all',
          docGroupFilter: 'all',
          docCategoryFilter: 'all',
          docGroupLevelFilter: 'all',
          docLeafCategoryFilter: 'all',
          docSubCategoryFilter: 'all',
          docSubfolderOtherFilter: 'all',
          catFilter: 'all',
          attachmentFilter: 'all',
          docListPage: 0,
          folderPathStack: rootStack,
          folderNavHistory: resetToRoot
            ? [{ folderPathStack: rootStack, docMainFolder: this.state.docMainFolder }]
            : this.state.folderNavHistory,
          folderNavIndex: resetToRoot ? 0 : this.state.folderNavIndex,
        });
      } else {
        this.setState({
          textFilter: '',
          vesselFilter: 'all',
          docGroupFilter: 'all',
          docCategoryFilter: 'all',
          docGroupLevelFilter: 'all',
          docLeafCategoryFilter: 'all',
          docSubCategoryFilter: 'all',
          catFilter: 'all',
          attachmentFilter: 'all',
          docListPage: 0,
          folderPathStack: [],
          docMainFolder: null,
          docScopeType: 'vessels',
          docViewMode: 'folder',
          folderNavHistory: [{ folderPathStack: [], docMainFolder: null }],
          folderNavIndex: 0,
        });
      }
    } else if (view === 'vessels') {
      // "Manage vessels →" on the Home dashboard sends the user here while
      // they have a specific SharePoint site selected there
      // (dashboardSiteFilter, e.g. NissenKaiunExternal) — land on that same
      // site instead of whatever vesselSiteFilter was last left at (often
      // the default Communication Site, which has no vessel folders, so
      // this page looked empty even though the dashboard tile they just saw
      // said otherwise). Only do this when actually arriving FROM the
      // dashboard (previousView === 'dashboard'); plain sidebar navigation
      // into Vessels leaves vesselSiteFilter — and any site the user
      // deliberately picked on this page — alone.
      const incomingSiteFilter =
        previousView === 'dashboard' && this.state.dashboardSiteFilter && this.state.dashboardSiteFilter !== 'all'
          ? this.state.dashboardSiteFilter
          : this.state.vesselSiteFilter;
      this.setState({ panelLoading: true, vesselSiteFilter: incomingSiteFilter || 'all' });
      try {
        const vesselUrl = incomingSiteFilter && incomingSiteFilter !== 'all'
          ? `${this._base()}/api/vessels?site_key=${encodeURIComponent(incomingSiteFilter)}`
          : `${this._base()}/api/vessels`;
        const data = await this._fetchJson(vesselUrl);
        if (data && Array.isArray(data)) {
          const { spoDeletedVesselIds, recycleBin } = this.state;
          // Filter out vessels that were soft-deleted via SPO delta sync
          // (backend DELETE may still be in-flight or the DB may not have
          // updated yet) — see _filterDeletedVessels for how live
          // SharePoint-discovered rows are treated differently.
          this.setState({
            vessels: this._filterDeletedVessels(
              data.map((v: any) => ({ ...v, name: cleanName(v.name), status: v.status || 'Active' })),
              recycleBin,
              spoDeletedVesselIds,
              incomingSiteFilter || null,
            ),
            panelLoading: false,
          });
        } else {
          this.setState({ panelLoading: false });
        }
      } catch {
        this.setState({ panelLoading: false });
      }
    } else if (view === 'recycle') {
      void this._loadRecycleBin();
    } else if (view === 'archive') {
      await this._loadArchiveList();
    } else if (view === 'users') {
      this.setState({ panelLoading: true });
      try {
        const data = await this._fetchJson(`${this._base()}/api/users`);
        // Backend's real role model (db/models.py UserProfile.role +
        // services/authorization.py) is Admin/User only — the old mock
        // roster's 4-tier scheme (Administrator/Manager/Reviewer/User) no
        // longer matches what /api/users returns ("Admin"/"User"), so
        // 'admin' used to fall through to the 'User' default and every
        // admin displayed as a plain user.
        const asRole = (value: any): UserItem['role'] => {
          return String(value || '').toLowerCase() === 'admin' ? 'Admin' : 'User';
        };
        const asStatus = (value: any): UserItem['status'] => {
          return String(value || '').toLowerCase() === 'active' ? 'Active' : 'Inactive';
        };
        const asPermissions = (value: any): UserItem['permissions'] => {
          if (!Array.isArray(value)) return [];
          return value.map((p: any) => ({
            site_key: String(p?.site_key || ''),
            can_view: Boolean(p?.can_view),
            can_upload: Boolean(p?.can_upload),
            can_tag_on_upload: Boolean(p?.can_tag_on_upload),
          })).filter(p => p.site_key);
        };
        const users: UserItem[] = Array.isArray(data)
          ? data.map((row: any, idx: number) => {
              const email = String(row?.email || '').trim().toLowerCase();
              const fallbackName = email ? email.split('@')[0] : `user-${idx + 1}`;
              return {
                id: String(row?.id || email || `user-${idx + 1}`),
                name: String(row?.name || fallbackName),
                email,
                role: asRole(row?.role),
                status: asStatus(row?.status),
                lastLogin: String(row?.lastLogin || 'Never'),
                permissions: asPermissions(row?.permissions),
              };
            })
          : [];
        this.setState({ usersList: users, panelLoading: false });
      } catch {
        this.setState({ usersList: [], panelLoading: false });
      }
    } else if (view === 'bento_email' || view === 'email_notify') {
      this.setState({ panelLoading: true, bentoLogs: [] });
      try {
        const data = await this._fetchJson(`${this._base()}/api/email-logs`);
        this.setState({ bentoLogs: data || [], panelLoading: false });
      } catch {
        this.setState({ bentoLogs: [], panelLoading: false });
      }
    }
  };

  /** Admin-only: purges every AI Bento Email log/attachment row via
   * DELETE /api/email-logs (server-gated by settings.admin_email_set —
   * a non-admin gets a 403 and the message below surfaces it). Clears
   * the grid to empty on success without touching the schema/columns. */
  public _clearAllBentoLogs = async (): Promise<void> => {
    if (this.state.bentoClearAllBusy) return;
    this.setState({ bentoClearAllBusy: true, bentoClearAllErr: null });
    try {
      await this._fetchJson(`${this._base()}/api/email-logs`, {
        method: 'DELETE',
        headers: this._headers(),
      });
      this.setState({ bentoLogs: [], bentoClearAllBusy: false, bentoClearAllErr: null });
    } catch (e: any) {
      this.setState({
        bentoClearAllBusy: false,
        bentoClearAllErr: e?.message || 'Failed to clear AI Bento Email logs.',
      });
    }
  };

  public _refreshCurrentModule = async (): Promise<void> => {
    const currentView = this.state.view;
    this.setState(prev => ({ reloadKey: prev.reloadKey + 1 }));
    if (currentView === 'dashboard') {
      await this._loadDashboardStats(true);
      return;
    }
    if (currentView === 'list' || currentView === 'vessels') {
      await this._loadData(true);
      return;
    }
    await this._goToView(currentView);
  };

  // ── CRUD Handlers ─────────────────────────────────────────────────────────

   public _openCreate = (prefillName?: string): void => {
    this.setState({ modal: 'create', selectedVessel: null, form: { ...BLANK_FORM, name: prefillName || '' }, modalMsg: null, modalError: null, formFieldErrors: {} });
  };

  public _openEditVessel = (v: VesselRecord): void => {
    this.setState({
      modal: 'edit',
      vesselActionPicker: null,
      selectedVessel: v,
      form: {
        name: v.name, imo: v.imo || '',
        shipyard: v.shipyard || '', hull_number: v.hull_number || '',
        vessel_type: v.vessel_type || '',
        target_site_ids: v.provisioned_site_ids || [],
      },
      modalMsg: null, modalError: null, formFieldErrors: {},
    });
  };

  public _openDeleteVessel = (v?: VesselRecord): void => {
    this._clearDeleteAutoCloseTimer();
    const selected = v || this.state.selectedVessel;
    this.setState({
      modal: 'delete',
      vesselActionPicker: null,
      selectedVessel: selected || null,
      deleteVesselIds: selected ? new Set([selected.id]) : new Set(),
      deleteVesselProgress: {},
      modalMsg: null,
      modalError: null,
      deleteAutoCloseSeconds: null,
    });
  };

  public _openVesselActionPicker = (action: 'edit' | 'delete'): void => {
    this._clearDeleteAutoCloseTimer();
    if (action === 'delete') {
      const selectedIds = this.state.selectedVessel ? new Set([this.state.selectedVessel.id]) : new Set<string>();
      this.setState({ vesselActionPicker: action, modalError: null, deleteVesselIds: selectedIds, deleteVesselProgress: {}, deleteAutoCloseSeconds: null });
      return;
    }
    this.setState({ vesselActionPicker: action, modalError: null, deleteAutoCloseSeconds: null });
  };

  // Stable per-card key for a not-yet-confirmed "Found in SharePoint" vessel
  // (it has no DB id yet), used to key discoveredVesselDrafts / confirmingVesselKey.
  public _discoveredVesselKey = (v: VesselRecord): string => {
    return `${v.provisioned_site_key || ''}::${v.vessel_folder_path || v.name}`;
  };

  public _updateDiscoveredVesselDraft = (
    v: VesselRecord,
    field: 'imo' | 'hull_number',
    value: string,
  ): void => {
    const key = this._discoveredVesselKey(v);
    this.setState(prev => {
      const existing = prev.discoveredVesselDrafts[key] || { imo: v.imo || '', hull_number: v.hull_number || '' };
      return {
        discoveredVesselDrafts: {
          ...prev.discoveredVesselDrafts,
          [key]: { ...existing, [field]: value },
        },
      };
    });
  };

  // Local (not-yet-saved) typing state for a regular DB vessel card's
  // IMO / Hull No. / Shipyard / Type field — offered only while that field
  // is empty (see renderVesselsPage). Committed to the database by
  // _saveVesselField, unlike discoveredVesselDrafts which waits for an
  // explicit "Confirm" click.
  public _updateVesselFieldDraft = (
    v: VesselRecord,
    field: 'imo' | 'hull_number' | 'shipyard' | 'vessel_type',
    value: string,
  ): void => {
    this.setState(prev => ({
      vesselFieldDrafts: {
        ...prev.vesselFieldDrafts,
        [v.id]: { ...(prev.vesselFieldDrafts[v.id] || {}), [field]: value },
      },
    }));
  };

  // Human label for the one-time-save confirmation dialog.
  private static readonly VESSEL_FIELD_LABELS: Record<string, string> = {
    imo: 'IMO',
    hull_number: 'Hull Number',
    shipyard: 'Shipyard',
    vessel_type: 'Type',
  };

  // Called on blur/select instead of saving straight away. A vessel-card
  // field is only ever offered once (while it's still empty — see
  // renderVesselsPage/renderEditableField), and it is locked immediately
  // after a successful save, so this is the one and only chance to write
  // it. Opens the confirmation dialog rather than PATCHing directly;
  // no-ops on an unchanged/empty value or while a save for the same field
  // is already in flight.
  public _requestSaveVesselField = (
    v: VesselRecord,
    field: 'imo' | 'hull_number' | 'shipyard' | 'vessel_type',
    rawValue: string,
  ): void => {
    const value = (rawValue || '').trim();
    if (!value || value === ((v as any)[field] || '')) return;
    const fieldKey = `${v.id}:${field}`;
    if (this.state.vesselFieldSaving[fieldKey]) return;
    this.setState({
      vesselFieldConfirm: {
        vesselId: v.id,
        vesselName: v.name,
        field,
        fieldLabel: VesselEmail.VESSEL_FIELD_LABELS[field] || field,
        value,
      },
    });
  };

  // Confirm button on the vessel-field confirmation dialog: commits the
  // pending edit (dialog state) via _saveVesselField, then closes it.
  public _confirmSaveVesselField = (): void => {
    const pending = this.state.vesselFieldConfirm;
    if (!pending) return;
    const v = this.state.vessels.find(row => row.id === pending.vesselId);
    this.setState({ vesselFieldConfirm: null });
    if (!v) return;
    this._saveVesselField(v, pending.field, pending.value).catch(() => undefined);
  };

  // Cancel/close on the vessel-field confirmation dialog: discards the
  // pending edit and reverts the field back to its (still empty) draft.
  public _cancelSaveVesselField = (): void => {
    const pending = this.state.vesselFieldConfirm;
    if (!pending) return;
    this.setState(prev => {
      const drafts = { ...prev.vesselFieldDrafts };
      if (drafts[pending.vesselId]) {
        const d = { ...drafts[pending.vesselId] };
        delete (d as any)[pending.field];
        drafts[pending.vesselId] = d;
      }
      return { vesselFieldConfirm: null, vesselFieldDrafts: drafts };
    });
  };

  // PATCH /api/vessels/{id} — saves a single IMO/Hull Number/Shipyard/Type
  // value the user filled in on a vessel card, after they've confirmed it
  // via the dialog (_requestSaveVesselField / _confirmSaveVesselField).
  // Only ever called for a field that was empty (the card only renders the
  // input in that case), and only touches the one field named —
  // _execute_update_vessel only renames SharePoint folders when `name`
  // itself changes, so this is a plain, side-effect-free DB write. No-ops
  // on an unchanged/empty value, and ignores a second save already in
  // flight for the same field. Once this succeeds, the field's saved value
  // makes renderEditableField render it as locked, read-only text instead
  // of an input — a vessel-card field can be saved exactly once.
  public _saveVesselField = async (
    v: VesselRecord,
    field: 'imo' | 'hull_number' | 'shipyard' | 'vessel_type',
    rawValue: string,
  ): Promise<void> => {
    const value = (rawValue || '').trim();
    if (!value || value === ((v as any)[field] || '')) return;
    const fieldKey = `${v.id}:${field}`;
    if (this.state.vesselFieldSaving[fieldKey]) return;
    this.setState(prev => ({
      vesselFieldSaving: { ...prev.vesselFieldSaving, [fieldKey]: true },
      vesselFieldError: { ...prev.vesselFieldError, [fieldKey]: '' },
    }));
    try {
      const res = await fetch(`${this._base()}/api/vessels/${v.id}`, {
        method: 'PATCH',
        headers: this._headers(),
        body: JSON.stringify({ [field]: value }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.detail || data?.message || `Save failed (${res.status})`);
      this.setState(prev => {
        const vessels = prev.vessels.map(row => row.id === v.id
          ? {
              ...row,
              imo: data.imo ?? row.imo,
              shipyard: data.shipyard ?? row.shipyard,
              hull_number: data.hull_number ?? row.hull_number,
              vessel_type: data.vessel_type ?? row.vessel_type,
            }
          : row);
        const drafts = { ...prev.vesselFieldDrafts };
        if (drafts[v.id]) {
          const d = { ...drafts[v.id] };
          delete (d as any)[field];
          drafts[v.id] = d;
        }
        return {
          vessels,
          vesselFieldDrafts: drafts,
          vesselFieldSaving: { ...prev.vesselFieldSaving, [fieldKey]: false },
        };
      });
    } catch (e: any) {
      this.setState(prev => ({
        vesselFieldSaving: { ...prev.vesselFieldSaving, [fieldKey]: false },
        vesselFieldError: { ...prev.vesselFieldError, [fieldKey]: e?.message || 'Save failed.' },
      }));
    }
  };

  // POST /api/vessels/sync-from-sharepoint — scans connected SharePoint sites'
  // root folders, reconciles against the DB + Term Store, and flags conflicts.
  // Runs across all sites unless a specific site is currently selected in the filter.
  // Called from _loadData on every vessel-list fetch, but throttled to at
  // most once per AUTO_SYNC_MIN_INTERVAL_MS per site_key so it doesn't
  // re-scan SharePoint on every reload. Silent — errors are swallowed here
  // (there's no manual "Sync" button anymore to surface them to), and the
  // vessel list is only reloaded if the sync actually found something new.
  // Drop vessels the user has just deleted from a freshly fetched vessel
  // list, and tag each remaining row with the site it was fetched for.
  //
  // Deleted rows are hidden by id (spoDeletedVesselIds) always, and by name
  // only when that name was deleted in the last RECENT_DELETE_GRACE_MS —
  // enough to cover a backend DELETE still in flight / the backend's short
  // scan cache right after an in-app delete. The backend never returns
  // deleted DB vessels or folders that no longer exist, so matching names
  // against the WHOLE Recycle Bin was wrong: old deleted records share names
  // with vessels that are active today, which is what hid 23 of NKSDocMan's
  // 25 vessels and 23 of NissenKaiunExternal's 28.
  //
  // fetchedForSite: the site_key this list was requested with
  // (/api/vessels?site_key=...). The backend already scoped the list using
  // the configured site registry (including sites that share a document
  // library under two keys), so VesselsPage trusts this tag instead of
  // re-matching each row against the partial site list the client knows.
  private static readonly RECENT_DELETE_GRACE_MS = 10 * 60 * 1000; // 10 minutes
  public _filterDeletedVessels = (
    list: any[],
    recycleBin: DeletedNode[],
    spoDeletedVesselIds: Set<string>,
    fetchedForSite?: string | null,
  ): any[] => {
    const now = Date.now();
    const deletedRecently = new Set<string>();
    for (const r of recycleBin || []) {
      if (r.kind !== 'vessel' && r.item_type !== 'vessel') continue;
      const key = (r.name || '').trim().toLowerCase();
      if (!key) continue;
      // Backend timestamps are naive UTC ("2026-09-18T17:00:08.75"); without
      // a zone suffix JS would parse them as local time, skewing the window.
      const rawDeletedAt = r.deleted_at ? String(r.deleted_at) : '';
      const deletedAt = rawDeletedAt
        ? Date.parse(/[zZ]|[+-]\d{2}:?\d{2}$/.test(rawDeletedAt) ? rawDeletedAt : `${rawDeletedAt}Z`)
        : NaN;
      if (!isNaN(deletedAt) && now - deletedAt < VesselEmail.RECENT_DELETE_GRACE_MS) {
        deletedRecently.add(key);
      }
    }
    // 'all' when the list was fetched without a site filter.
    const siteTag = fetchedForSite && fetchedForSite !== 'all' ? fetchedForSite : 'all';
    return list
      .filter((v: any) => !spoDeletedVesselIds.has(v.id) && !deletedRecently.has((v.name || '').trim().toLowerCase()))
      .map((v: any) => ({ ...v, _fetched_for_site: siteTag }));
  };

  public _maybeAutoSyncVesselsFromSharePoint = (siteKey: string | null): void => {
    const throttleKey = siteKey || 'all';
    const now = Date.now();
    const last = this._lastAutoSyncAt[throttleKey] || 0;
    if (now - last < VesselEmail.AUTO_SYNC_MIN_INTERVAL_MS) return;
    this._lastAutoSyncAt[throttleKey] = now;
    void (async () => {
      try {
        const res = await fetch(`${this._base()}/api/vessels/sync-from-sharepoint`, {
          method: 'POST',
          headers: this._headers(),
          body: JSON.stringify(siteKey ? { site_key: siteKey } : {}),
        });
        if (!res.ok) return;
        const data = await res.json().catch(() => ({}));
        const found = (data?.new ?? 0) > 0 || (data?.updated ?? 0) > 0;
        if (found && !this._isUnmounted) {
          await this._loadData(true);
        }
      } catch {
        /* best-effort background sync */
      }
    })();
  };

  // POST /api/vessels/confirm-discovered — turns a "Found in SharePoint" card
  // into a real DB-backed vessel record at its existing SharePoint folder path
  // (no new folder is created; see confirm_discovered_vessel in real_backend.py).
  public _confirmDiscoveredVessel = async (v: VesselRecord): Promise<void> => {
    const key = this._discoveredVesselKey(v);
    if (this.state.confirmingVesselKey) return;
    const draft = this.state.discoveredVesselDrafts[key] || { imo: v.imo || '', hull_number: v.hull_number || '' };
    const setCardError = (message: string | null): void => {
      this.setState(prev => {
        const errors = { ...prev.discoveredVesselConfirmError };
        if (message) errors[key] = message; else delete errors[key];
        return { discoveredVesselConfirmError: errors };
      });
    };
    if (!v.provisioned_site_key || !v.vessel_folder_path) {
      setCardError('Missing SharePoint site/path for this vessel — refresh the vessel list and try again.');
      return;
    }
    // The backend silently fills in a random placeholder IMO when this is
    // left blank (see _validate_vessel_input in real_backend.py), so a
    // click here used to "succeed" with a made-up IMO the user never typed
    // — which then never matched what they'd entered, making it look like
    // their input was never saved. Require a real 7-digit IMO up front
    // instead, and surface the problem right under the button rather than
    // only in the page-level banner above the vessel grid, which is easy
    // to miss once the user has scrolled down to this card.
    const imoValue = (draft.imo || '').trim();
    if (!/^\d{7}$/.test(imoValue)) {
      setCardError('Enter a valid 7-digit IMO number before confirming.');
      return;
    }
    setCardError(null);
    this.setState({ confirmingVesselKey: key, vesselSyncError: null });
    try {
      const res = await fetch(`${this._base()}/api/vessels/confirm-discovered`, {
        method: 'POST',
        headers: this._headers(),
        body: JSON.stringify({
          name: v.name,
          imo: imoValue,
          hull_number: draft.hull_number || null,
          site_key: v.provisioned_site_key,
          original_path: v.vessel_folder_path,
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (res.status === 409) {
        this.setState({ confirmingVesselKey: null });
        setCardError(data?.message || 'A vessel with this name already exists.');
        return;
      }
      if (!res.ok) throw new Error(data?.detail || data?.error || data?.message || `Confirm failed (${res.status})`);
      this.setState(prev => {
        const drafts = { ...prev.discoveredVesselDrafts };
        delete drafts[key];
        return { confirmingVesselKey: null, discoveredVesselDrafts: drafts };
      });
      await this._goToView('vessels');
    } catch (error: any) {
      this.setState({ confirmingVesselKey: null });
      setCardError(error?.message || 'Confirming this vessel failed.');
    }
  };

  public _closeModal = (): void => {
    this._clearDeleteAutoCloseTimer();
    if (!this.state.modalBusy) this.setState({ modal: 'none', vesselActionPicker: null, modalMsg: null, modalError: null, deleteVesselProgress: {}, deleteVesselReason: '', deleteAutoCloseSeconds: null });
  };

  public _submitCreate = async (): Promise<void> => {
    const { form, vessels } = this.state;
    const normalizedName = (value: string): string =>
      value.replace(/[ _'\"]+/g, '').toLowerCase();
    if (!form.name.trim()) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, name: 'Vessel name is required.' } });
      return;
    }
    if (vessels.some(v => normalizedName(v.name) === normalizedName(form.name))) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, name: 'Vessel name already exists.' } });
      return;
    }
    if (!form.imo.trim()) { this.setState({ modalError: 'IMO number is required.' }); return; }
    if (!/^\d{7}$/.test(form.imo.trim())) { this.setState({ modalError: 'IMO number must be exactly 7 digits.' }); return; }
    if (!form.site_key) { this.setState({ modalError: 'Select a SharePoint site.' }); return; }
    if (form.parent_folder_path === undefined) { this.setState({ modalError: 'Choose a parent folder in SharePoint.' }); return; }
    this.setState({ modalBusy: true, modalError: null, modalMsg: null });

    const defaultSiteKey = this.state.activeDocumentSite || form.site_key || this.state.documentSites[0]?.site_key;
    const newVesselRecord: VesselRecord = {
      id: `v_${Date.now()}`,
      name: form.name.trim(),
      imo: form.imo.trim(),
      shipyard: form.shipyard.trim() || undefined,
      hull_number: form.hull_number.trim() || undefined,
      vessel_type: form.vessel_type || undefined,
      status: 'Active',
      image_url: pickRandomVesselImage(form.vessel_type),
      provisioned_site_ids: form.site_key ? [form.site_key] : defaultSiteKey ? [defaultSiteKey] : [],
      provisioned_site_key: form.site_key || defaultSiteKey || undefined,
    };

        try {
      const res = await fetch(`${this._base()}/api/vessels`, {
        method: 'POST', headers: this._headers(),
        body: JSON.stringify({
          name: form.name.trim(),
          imo: form.imo.trim(),
          shipyard: form.shipyard.trim() || null,
          hull_number: form.hull_number.trim() || null,
          vessel_type: form.vessel_type || null,
          site_key: form.site_key,
          parent_folder_path: form.parent_folder_path,
          subfolders: form.subfolders || [],
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok && res.status !== 202) {
        const serverMsg: string = data?.detail || data?.message || `Error ${res.status}`;
        // Duplicate IMO (and similar validation) errors go under the IMO
        // field itself; anything else is a generic modal error. Either way,
        // STOP here — do not fall through to the fake "offline" success path.
        if (/vessel name|already exists/i.test(serverMsg)) {
          this.setState({ modalBusy: false, modalError: null, formFieldErrors: { ...this.state.formFieldErrors, name: serverMsg } });
        } else if (/imo/i.test(serverMsg)) {
          this.setState({ modalBusy: false, modalError: null, formFieldErrors: { ...this.state.formFieldErrors, imo: serverMsg } });
        } else {
          this.setState({ modalBusy: false, modalError: serverMsg });
        }
        return;
      }

      if (data.id || data.result?.id) {
        newVesselRecord.id = data.id || data.result.id;
      }

      // Add new vessel immediately to top of state grid
      this.setState(prev => ({
        vessels: [newVesselRecord, ...prev.vessels.filter(v => v.name.toLowerCase() !== newVesselRecord.name.toLowerCase())],
        // Clear any stale folder provision banner from a previous operation so a fresh "0 created"
        // result from the redundant front-end call does not appear. The backend already kicked off
        // provisioning via asyncio.create_task during POST /api/vessels — we do not need to call
        // _provisionVesselFolders here. The vessel state refreshes when the success modal closes.
        folderCreationResults: null,
        folderCreationError: null,
      }));

      // The vessel folder was just created directly in this site's drive
      // (POST /api/vessels awaits the SharePoint folder creation for a
      // custom site/parent path before responding). Drop the cached live
      // folder tree for that site so the Documents module shows the new
      // folder immediately instead of a stale pre-creation snapshot.
      this._invalidateDocumentLiveTree(form.site_key);

      // Transition modal to success screen immediately
      this.setState({
        modalBusy: false,
        modalMsg: `🎉 Vessel "${form.name}" created successfully!`,
        modalError: null,
        formFieldErrors: {},
      });
    } catch (e: any) {
      // A real network failure (fetch itself threw, server unreachable) —
      // NOT a validation rejection, that's handled above and returns early.
      this.setState(prev => ({
        vessels: [newVesselRecord, ...prev.vessels.filter(v => v.name.toLowerCase() !== newVesselRecord.name.toLowerCase())],
        modalBusy: false,
        modalMsg: `🎉 Vessel "${form.name}" created successfully!`,
        modalError: null,
        folderCreationResults: null,
        folderCreationError: null,
      }));
    }
  };

  public _submitEdit = async (): Promise<void> => {
    const { form, selectedVessel, vessels } = this.state;
    if (!selectedVessel) return;
    const normalizedName = (value: string): string =>
      value.replace(/[ _'\"]+/g, '').toLowerCase();
    const otherVessels = vessels.filter(v => v.id !== selectedVessel.id);
    if (!form.name.trim()) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, name: 'Vessel name is required.' } });
      return;
    }
    if (otherVessels.some(v => normalizedName(v.name) === normalizedName(form.name))) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, name: 'Vessel name already exists.' } });
      return;
    }
    if (!form.imo.trim()) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, imo: 'IMO number is required.' } });
      return;
    }
    if (!/^\d{7}$/.test(form.imo.trim())) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, imo: 'IMO number must be exactly 7 digits.' } });
      return;
    }
    if (otherVessels.some(v => (v.imo || '').trim() === form.imo.trim())) {
      this.setState({ modalError: null, formFieldErrors: { ...this.state.formFieldErrors, imo: 'A vessel with that IMO number already exists.' } });
      return;
    }
    this.setState({ modalBusy: true, modalError: null, modalMsg: null });
    try {
      const res = await fetch(`${this._base()}/api/vessels/${selectedVessel.id}`, {
        method: 'PATCH', headers: this._headers(),
        body: JSON.stringify({
          name: form.name.trim() || null,
          imo: form.imo.trim() || null,
          shipyard: form.shipyard.trim() || null,
          hull_number: form.hull_number.trim() || null,
          vessel_type: form.vessel_type || null,
          provisioned_site_ids: form.target_site_ids && form.target_site_ids.length > 0 ? form.target_site_ids : undefined,
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok && res.status !== 202) {
        const serverMsg: string = data?.detail || data?.message || `Error ${res.status}`;
        if (/vessel name|name.*exists/i.test(serverMsg)) {
          this.setState({ modalBusy: false, modalError: null, formFieldErrors: { ...this.state.formFieldErrors, name: serverMsg } });
        } else if (/imo/i.test(serverMsg)) {
          this.setState({ modalBusy: false, modalError: null, formFieldErrors: { ...this.state.formFieldErrors, imo: serverMsg } });
        } else {
          this.setState({ modalBusy: false, modalError: serverMsg });
        }
        return;
      }
      const msg = data.status === 'pending' ? '⏳ Vessel update submitted for approval.' : `✅ Vessel updated successfully.`;
      this.setState({ modalBusy: false, modalMsg: msg });
    } catch (e: any) {
      const updated = vessels.map(v => v.id === selectedVessel.id ? { ...v, name: form.name.trim() || v.name, imo: form.imo.trim() || v.imo, shipyard: form.shipyard || v.shipyard, hull_number: form.hull_number || v.hull_number, vessel_type: form.vessel_type || v.vessel_type } : v);
      this.setState({ vessels: updated, modalBusy: false, modalMsg: '✅ Vessel updated successfully.' });
    }
  };

  public _submitDelete = async (): Promise<void> => {
    const { deleteVesselIds, vessels, deleteVesselReason } = this.state;
    const selected = vessels.filter(v => deleteVesselIds.has(v.id));
    if (selected.length === 0) {
      this.setState({ modalError: 'Select at least one vessel to delete.' });
      return;
    }
    this._clearDeleteAutoCloseTimer();
    const progressSeed: Record<string, { name: string; status: 'waiting' | 'deleting' | 'success' | 'pending' | 'failed'; message?: string }> = {};
    selected.forEach(v => { progressSeed[v.id] = { name: v.name, status: 'waiting' }; });

    this.setState({ modalBusy: true, modalError: null, deleteAutoCloseSeconds: null, deleteVesselProgress: progressSeed });
    const deletedIds: string[] = [];
    const pendingNames: string[] = [];
    const failures: string[] = [];

    for (const vessel of selected) {
      this.setState(prev => ({
        deleteVesselProgress: {
          ...prev.deleteVesselProgress,
          [vessel.id]: { ...(prev.deleteVesselProgress[vessel.id] || { name: vessel.name }), status: 'deleting' },
        },
      }));

      // Backend is the single source of truth for vessel deletion (SPO move +
      // DB cleanup + recycle-bin bookkeeping). Avoid client-side pre-delete to
      // prevent race conditions where backend sees folders as already missing.
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 45000);
      try {
        const reasonParam = deleteVesselReason.trim() ? `&reason=${encodeURIComponent(deleteVesselReason.trim())}` : '';
        const res = await fetch(`${this._base()}/api/vessels/${vessel.id}?vessel_name=${encodeURIComponent(vessel.name)}${reasonParam}`, {
          method: 'DELETE', headers: this._headers(), signal: controller.signal,
        });
        const raw = await res.text();
        let data: any = {};
        try { data = raw ? JSON.parse(raw) : {}; } catch { /* response is not JSON */ }
        if (!res.ok && res.status !== 202) throw new Error(data?.message || raw || `Error ${res.status}`);
        if (data.status === 'pending' || res.status === 202) {
          pendingNames.push(vessel.name);
          this.setState(prev => ({
            deleteVesselProgress: {
              ...prev.deleteVesselProgress,
              [vessel.id]: {
                ...(prev.deleteVesselProgress[vessel.id] || { name: vessel.name }),
                status: 'pending',
                message: data?.message || 'Submitted for approval.',
              },
            },
          }));
        } else {
          deletedIds.push(vessel.id);
          this.setState(prev => ({
            deleteVesselProgress: {
              ...prev.deleteVesselProgress,
              [vessel.id]: {
                ...(prev.deleteVesselProgress[vessel.id] || { name: vessel.name }),
                status: 'success',
                message: data?.message || 'Moved to Recycle Bin.',
              },
            },
          }));
        }
      } catch (e: any) {
        const timedOut = e?.name === 'AbortError';
        failures.push(`${vessel.name}: ${timedOut ? 'Delete request timed out. Please retry.' : (e?.message || 'Delete failed')}`);
        this.setState(prev => ({
          deleteVesselProgress: {
            ...prev.deleteVesselProgress,
            [vessel.id]: {
              ...(prev.deleteVesselProgress[vessel.id] || { name: vessel.name }),
              status: 'failed',
              message: timedOut ? 'Delete request timed out. Please retry.' : (e?.message || 'Delete failed'),
            },
          },
        }));
      } finally {
        clearTimeout(timeout);
      }
    }

    const removedCount = deletedIds.length;
    const deletedNamesNorm = new Set(
      selected
        .filter(v => deletedIds.includes(v.id))
        .map(v => cleanName(v.name).trim().toLowerCase())
    );
    const messageParts: string[] = [];
    if (removedCount) messageParts.push(`${removedCount} vessel${removedCount === 1 ? '' : 's'} moved to Recycle Bin.`);
    if (pendingNames.length) messageParts.push(`${pendingNames.length} deletion request${pendingNames.length === 1 ? '' : 's'} submitted for approval.`);

    // Build immediate recycle bin entries for deleted vessels so they appear
    // in the Recycle Bin page right away without waiting for the backend poll.
    const now = new Date().toISOString();
    const activeSiteObj = (this.state.documentSites || []).find(s => s.site_key === this.state.activeDocumentSite);
    const immediateRecycleBinEntries: DeletedNode[] = selected
      .filter(v => deletedIds.includes(v.id))
      .map(v => {
        const siteNames = (v.provisioned_site_ids || []).map(sk => {
          const matched = (this.state.documentSites || []).find(s => s.site_key === sk);
          return matched?.sp_site_name || sk;
        }).filter(Boolean);
        const resolvedSiteName = siteNames.length > 0
          ? siteNames.join(', ')
          : (activeSiteObj?.sp_site_name || this.props.siteUrl || 'SharePoint');

        return {
          id: v.id,
          name: v.name,
          kind: 'vessel' as const,
          item_type: 'vessel',
          main_folder: 'Vessels',
          original_path: `Vessels/Specific Vessels/${v.name}`,
          vessel_name: '',
          category: '',
          sub_category: '',
          deleted_at: now,
          vessel_imo: v.imo,
          vessel_type: v.vessel_type,
          site_name: resolvedSiteName,
          site_key: (v.provisioned_site_ids || [])[0] || this.state.activeDocumentSite || undefined,
        };
      });

    this.setState(prev => ({
      modalBusy: false,
      modalMsg: messageParts.join(' '),
      modalError: failures.length ? failures.join(' | ') : null,
      vessels: prev.vessels.filter(v => !deletedIds.includes(v.id)),
      vesselFilter: deletedNamesNorm.has(cleanName(prev.vesselFilter || '').trim().toLowerCase()) ? 'all' : prev.vesselFilter,
      selectedVessel: deletedIds.includes(prev.selectedVessel?.id || '') ? null : prev.selectedVessel,
      deleteVesselIds: new Set(pendingNames.length || failures.length ? selected.filter(v => !deletedIds.includes(v.id)).map(v => v.id) : []),
      // Add deleted vessels to recycle bin immediately (deduplicated by name)
      recycleBin: [
        ...immediateRecycleBinEntries,
        ...prev.recycleBin.filter(r => !immediateRecycleBinEntries.some(e => e.name.toLowerCase() === r.name.toLowerCase())),
      ],
    }));

    if (removedCount > 0) {
      // Trigger an immediate delta sync so the deleted folders are pruned from the local folder map
      void this._syncScheduler?.triggerNow().catch(() => undefined);

      const deletedVessels = selected.filter(v => deletedIds.includes(v.id));
      this.setState(prev => ({
        alertsList: [
          ...deletedVessels.map(v => ({
            id: `vessel_deleted_${v.id}_${Date.now()}`,
            drive_item_id: v.id,
            folder_name: v.name,
            folder_path: `Vessels/Specific Vessels/${v.name}`,
            parent_folder_id: null,
            vessel_name: v.name,
            department: 'All Departments',
            created_by_email: this.props.userEmail || '',
            created_by_name: this.props.userEmail?.split('@')[0] || 'Admin',
            alert_type: 'vessel_deleted' as any,
            read: false,
            created_at: now,
          })),
          ...prev.alertsList,
        ],
      }));
    }

    // Only after full backend success (no pending approvals/failures),
    // start a visible 10-second auto-close countdown in the success popup.
    if (removedCount > 0 && pendingNames.length === 0 && failures.length === 0) {
      this._startDeleteAutoClose(10);
    }

    // Re-read the server recycle bin so restore/permanent-delete actions use real item ids.
    // Poll with retries because the SharePoint recycle bin API has a short propagation delay
    // after a delete — items deleted milliseconds ago may not appear on the first fetch.
    if (removedCount) {
      const fetchRecycleBin = async (attempt: number): Promise<void> => {
        try {
          const data = await this._fetchJson(`${this._base()}/api/recycle-bin/nodes`);
          const nodes = (data || []).map((item: any) => ({ ...item, name: cleanName(item.name || '') }));
          // If we deleted N vessels but got fewer back, retry once after a short delay
          if (nodes.filter((n: any) => n.kind === 'vessel').length < removedCount && attempt < 3) {
            setTimeout(() => fetchRecycleBin(attempt + 1), 1500);
          } else {
            this.setState(prev => {
              const backendIds = new Set(nodes.map((n: any) => n.id));
              const localItems = prev.recycleBin.filter(item => !backendIds.has(item.id));
              return { recycleBin: [...nodes, ...localItems] };
            });
          }
        } catch { /* The item will be visible after the next Recycle Bin refresh. */ }
      };
      void fetchRecycleBin(1);
    }
  };

  // SharePoint Folder Provisioning ─────────────────────────────────────────

  /**
   * Refreshes just-provisioned folder rows for ONE vessel via the fast,
   * per-vessel flat-tree DB query (backend/app/main.py list_vessels_flat_tree,
   * the `vessel_name` branch — a targeted query, not the full-fleet scan).
   *
   * Replaces the previous post-provision `rows: []` + `_loadData(true)`,
   * which reset the whole Documents list and forced a full reload across
   * EVERY vessel. When the DB-backed flat-tree briefly returns nothing for
   * a brand-new vessel (provisioning hasn't finished indexing yet), `_loadData`
   * falls through its Graph fallback straight to `_flattenAll`, which walks
   * every vessel's every main-folder and every category with its own
   * `/api/folders/{id}/children` call — i.e. exactly the "creating one vessel
   * triggers a burst/looping of children calls" symptom. Fetching only the
   * new vessel's rows here avoids that fan-out entirely; the full multi-vessel
   * fallback walk is never invoked as a side effect of provisioning.
   */
  public async _refreshSingleVesselRows(vesselName: string): Promise<void> {
    const base = this._base();
    if (!base) return;
    try {
      const url = `${base}/api/vessels/flat-tree?vessel_name=${encodeURIComponent(vesselName)}&force_refresh=true`;
      const data = await this._fetchJson(url);
      const freshRows = Array.isArray(data) ? this._normalize(data) : [];
      const targetName = cleanName(vesselName).trim().toLowerCase();
      this.setState(prev => {
        const keptRows = prev.rows.filter(r => cleanName(r.vesselName).trim().toLowerCase() !== targetName);
        const merged = this._mergeWithUploadCache([...keptRows, ...freshRows]);
        return { rows: merged.rows, uploadedFilesByFolder: merged.uploadedFilesByFolder };
      });
      void this._loadDashboardStats();
    } catch {
      // Targeted refresh failed (e.g. transient network error) — fall back
      // to the previous full-reload behavior rather than leaving stale rows.
      await this._loadData(true);
    }
  }

  public async _provisionVesselFolders(
    vesselName: string,
    vesselId?: string,
    targetSiteIds?: string[],
  ): Promise<{ success: boolean; results: FolderResult[] }> {
    // Determine which sites to provision on:
    // • If caller passes explicit targetSiteIds (e.g. from Create Vessel form), honour exactly those.
    // • Otherwise fall back to the currently active document site (legacy path for manual Provision button).
    const explicitSites = targetSiteIds && targetSiteIds.length > 0 ? targetSiteIds : null;
    const selectedSite = explicitSites ? explicitSites[0] : this.state.activeDocumentSite;
    if (explicitSites && vesselId && /^\d+$/.test(vesselId)) {
      // Multi-site explicit provision: call provision-sites with all chosen site keys
      this.setState({ folderCreationBusy: true, folderCreationResults: null, folderCreationError: null, folderProvisioningVesselId: vesselId });
      try {
        const response = await fetch(`${this._base()}/api/vessels/${encodeURIComponent(vesselId)}/provision-sites`, {
          method: 'POST',
          headers: this._headers(),
          body: JSON.stringify({ site_keys: explicitSites }),
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(data?.detail || data?.error || `Provisioning failed (${response.status})`);
        const allSucceeded = explicitSites.every(sk => {
          const sr = data?.results?.[sk] || {};
          return sr.status === 'success' || sr.success === true;
        });
        if (!allSucceeded) {
          const failedSites = explicitSites.filter(sk => {
            const sr = data?.results?.[sk] || {};
            return sr.status !== 'success' && sr.success !== true;
          });
          throw new Error(`Provisioning failed for sites: ${failedSites.join(', ')}`);
        }
        this.setState({
          folderCreationBusy: false,
          folderProvisioningVesselId: null,
          folderCreationError: null,
          folderCreationResults: null,
          provisionedVesselIds: new Set(Array.from(this.state.provisionedVesselIds).concat(vesselId)),
        });
        await this._refreshSingleVesselRows(vesselName);
        return { success: true, results: [] };
      } catch (error: any) {
        const message = error?.message || 'Folder provisioning failed.';
        this.setState({ folderCreationBusy: false, folderProvisioningVesselId: null, folderCreationResults: null, folderCreationError: message });
        return { success: false, results: [] };
      }
    }
    if (selectedSite && vesselId && /^\d+$/.test(vesselId)) {
      this.setState({ folderCreationBusy: true, folderCreationResults: null, folderCreationError: null, folderProvisioningVesselId: vesselId });
      try {
        const response = await fetch(`${this._base()}/api/vessels/${encodeURIComponent(vesselId)}/provision-sites`, {
          method: 'POST',
          headers: this._headers(),
          body: JSON.stringify({ site_keys: [selectedSite] }),
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(data?.detail || data?.error || `Provisioning failed (${response.status})`);
        const siteResult = data?.results?.[selectedSite] || {};
        const success = siteResult.status === 'success' || siteResult.success === true;
        if (!success) throw new Error(siteResult.error || `Provisioning failed for site ${selectedSite}`);
        this.setState({
          folderCreationBusy: false,
          folderProvisioningVesselId: null,
          folderCreationError: null,
          folderCreationResults: null,
          provisionedVesselIds: new Set(Array.from(this.state.provisionedVesselIds).concat(vesselId)),
        });
        await this._refreshSingleVesselRows(vesselName);
        return { success: true, results: [] };
      } catch (error: any) {
        const message = error?.message || 'Folder provisioning failed.';
        this.setState({ folderCreationBusy: false, folderProvisioningVesselId: null, folderCreationResults: null, folderCreationError: message });
        return { success: false, results: [] };
      }
    }

    // Every real caller provides a numeric vesselId (a genuine DB-backed
    // vessel always has one), so the branches above always handle normal
    // provisioning via the backend's flat-root /provision-sites endpoint.
    // Reaching this point means we were called without a valid vesselId
    // and/or without a resolvable SharePoint site — there is nothing left
    // to fall back to. This used to silently drop into a client-side Graph
    // builder (graphFolderService.ts's createVesselFolders(), now retired)
    // that recreated the old multi-department nested folder tree — wrong
    // for the current flat-root design, and it could write unwanted
    // folders straight into the customer's SharePoint. Surface a clear,
    // actionable error instead.
    const missing = !vesselId
      ? 'no vessel ID was provided'
      : !/^\d+$/.test(vesselId)
        ? `vessel ID "${vesselId}" is not a valid database ID`
        : 'no SharePoint site is selected';
    console.error(`[VesselDMS] _provisionVesselFolders: cannot provision — ${missing}. vesselName="${vesselName}" vesselId="${vesselId}" selectedSite="${selectedSite}"`);
    const message = `Could not provision folders for "${vesselName}": ${missing}. Please ensure the vessel was created successfully and a SharePoint site is selected, then try again.`;
    this.setState(prev => ({
      folderCreationBusy: false,
      folderProvisioningVesselId: null,
      folderCreationError: message,
      spoProvisionDialog: prev.spoProvisionDialog
        ? { ...prev.spoProvisionDialog, provisioning: false, done: true, error: message }
        : prev.spoProvisionDialog,
    }));
    return { success: false, results: [] };
  }

  // ── Load Files for Vessel ─────────────────────────────────────────────────

  /** Find the Documents list row represented by a live SharePoint folder ID. */
  public _getGroupKeyForLiveFolderId(folderId: string): string {
    const node = this.state.spoFolderMap.get(folderId);
    if (!node?.isFolder || !node.serverRelativePath) return '';

    const normalisePath = (path: string): string =>
      (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
    const livePath = normalisePath(node.serverRelativePath);
    const row = this.state.rows.find(candidate => {
      const expectedPath = normalisePath(this._sharePointFolderPath(candidate.subFolderPath, ''));
      return Boolean(expectedPath) &&
        (livePath === expectedPath || livePath.endsWith(`/${expectedPath}`));
    });
    return row?.groupKey || '';
  }

  // Debounced entry point for the Documents search box (see
  // DocumentsPage.tsx's textFilter onChange). The box itself only filters
  // `rows` already in state, which are loaded lazily per-vessel — so a
  // term matching a vessel the user hasn't opened yet would otherwise find
  // nothing. This asks the backend /api/search (which now also matches
  // Folder.path, i.e. group/category/vessel segments, not just the leaf
  // name — see real_backend.py/store.py search()) which vessels contain a
  // match, and lazily loads those vessels' rows so the existing local
  // filter picks them up on the next render.
  public _scheduleGlobalSearch = (term: string): void => {
    if (this._globalSearchTimer) {
      clearTimeout(this._globalSearchTimer);
      this._globalSearchTimer = null;
    }
    const trimmed = (term || '').trim();
    // Same floor as the field's own tokenized matching (a 1-char query is
    // rarely useful and would fan out a request per keystroke); the local
    // filter still runs immediately regardless of this floor.
    if (trimmed.length < 2) return;
    this._globalSearchTimer = setTimeout(() => {
      void this._triggerGlobalSearch(trimmed).catch(() => undefined);
    }, 400);
  };

  public async _triggerGlobalSearch(term: string): Promise<void> {
    const activeSiteKey = (this.state.activeDocumentSite || '').trim().toLowerCase();
    const key = `${activeSiteKey}|${this.state.docScopeType}|${term.toLowerCase()}`;
    // Once a term has been resolved to a set of vessels this session, it
    // never needs to hit the backend again — the matching vessels' rows
    // stay loaded (_filesLoadedForVessels), and the box's own local filter
    // handles re-filtering as the user keeps typing/editing the term.
    if (this._globalSearchSeenTerms.has(key)) return;

    const { vesselFilter, vessels } = this.state;
    const selectedSiteKey = (this.state.activeDocumentSite || '').trim();
    // Whether the just-switched-to site's dashboard scan hadn't finished
    // even once yet (see get_dashboard_documents on the backend) — an empty
    // `items` response in that state is a cold-cache placeholder, not a
    // real "no matches" answer. Marking `key` as seen below is skipped
    // while this is true, and the search is retried shortly instead, so a
    // term searched right after switching sites (before its background scan
    // completes) doesn't get stuck showing no results for the rest of the
    // session even once the real data lands — this is what made "search
    // right after switching the SharePoint site dropdown" look broken
    // while the same search worked fine once the site had been browsed a
    // little (e.g. via the Sites tile list) and its scan had time to finish.
    let sitePending = false;
    if (selectedSiteKey && selectedSiteKey !== 'all') {
      const params = new URLSearchParams({
        site_key: selectedSiteKey,
        q: term,
        page: '1',
        page_size: '200',
      });
      const siteSearch = await this._fetchJson(
        `${this._base()}/api/dashboard/documents?${params.toString()}`
      ).catch(() => null);
      sitePending = siteSearch?.scan_pending === true;
      const siteItems = Array.isArray(siteSearch?.items) ? siteSearch.items : [];
      if (siteItems.length > 0) {
        const site = this.state.documentSites.find(item => item.site_key === selectedSiteKey);
        const siteName = site?.sp_site_name || site?.site_key || selectedSiteKey;
        const libraryName = site?.default_library_name || 'Shared Documents';
        const searchRows: FlatRow[] = siteItems.map((item: any, index: number) => {
          const relativePath = String(item.subFolderPath || item.name || '').trim();
          const path = `SharePoint Sites > ${siteName} > ${libraryName} > ${relativePath}`;
          const pathParts = relativePath.split('>').map((part: string) => part.trim()).filter(Boolean);
          return {
            srNo: `site-search-${index}`,
            vesselName: item.vessel || 'Not Listed',
            group: pathParts[0] || libraryName,
            category: pathParts[1] || '',
            subCategory: pathParts[2] || '',
            subFolderPath: path,
            fileName: item.name || null,
            fileId: item.id || null,
            fileSize: item.fileSize,
            fileUploadedAt: item.createdEpoch || item.modifiedEpoch,
            canUpload: false,
            groupKey: `live:${selectedSiteKey}:search`,
            uploadFolderId: '',
            monthDriven: false,
            documentSection: pathParts[1] || '',
            tags: {},
            siteKey: selectedSiteKey,
          } as FlatRow;
        });
        this.setState(previous => {
          const existingIds = new Set(previous.rows.map(row => row.fileId).filter(Boolean));
          const newRows = searchRows.filter(row => row.fileId && !existingIds.has(row.fileId));
          return newRows.length > 0 ? { rows: [...previous.rows, ...newRows] } : null as any;
        });
        this._scheduleForceUpdate();
      }
    }

    if (sitePending) {
      // The real per-site scan is still running in the background —
      // try this exact term again shortly instead of leaving it marked
      // "seen" with a stale, placeholder-empty result for the rest of the
      // session. Not added to _globalSearchSeenTerms, so this call is
      // itself safe to repeat.
      window.setTimeout(() => {
        void this._triggerGlobalSearch(term).catch(() => undefined);
      }, 3000);
    } else {
      this._globalSearchSeenTerms.add(key);
    }

    let url = `${this._base()}/api/search?q=${encodeURIComponent(term)}`;
    if (vesselFilter && vesselFilter !== 'all') {
      const scopedVessel = vessels.find(v => v.name === vesselFilter);
      if (scopedVessel?.id) url += `&vessel_id=${encodeURIComponent(scopedVessel.id)}`;
    }

    const data = await this._fetchJson(url).catch(() => null);
    if (!Array.isArray(data) || data.length === 0) return;

    // Folder.path is "<group>/<vessel>/<category>/…", and _trail() (backend)
    // returns one {id,name} per path segment — so trail[1] is the vessel
    // folder name whenever a result sits under one (trail.length <= 1 means
    // the match is a top-level group/main folder itself, not vessel-scoped).
    const vesselNames = new Set<string>();
    data.forEach((item: any) => {
      const trail = Array.isArray(item.trail) ? item.trail : [];
      const vesselSeg = trail.length > 1 ? trail[1]?.name : null;
      if (vesselSeg) vesselNames.add(vesselSeg);
    });

    let loadedAny = false;
    vesselNames.forEach(name => {
      if (!this._filesLoadedForVessels.has(name)) {
        this._filesLoadedForVessels.add(name);
        loadedAny = true;
        void this._loadFilesForVessel(name).catch(() => undefined);
      }
    });
    if (loadedAny) this._scheduleForceUpdate();
  }

  public async _loadFilesForVessel(vesselName: string): Promise<void> {
    const { rows } = this.state;
    const vesselRows = rows.filter(r => r.vesselName === vesselName && r.uploadFolderId);
    const normalisePath = (path: string): string =>
      (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
    const liveFolderIdsByPath = new Map<string, string>();
    this.state.spoFolderMap.forEach(node => {
      if (!node.deleted && node.isFolder && node.serverRelativePath) {
        liveFolderIdsByPath.set(normalisePath(node.serverRelativePath), node.id);
      }
    });

    // DB folder IDs may belong to a previous drive or a deleted folder. Match
    // the breadcrumb to the live delta map before making a Graph request.
    const liveFolders = vesselRows.map(row => {
      const expectedPath = normalisePath(this._sharePointFolderPath(row.subFolderPath, ''));
      let id = liveFolderIdsByPath.get(expectedPath);
      if (!id && expectedPath) {
        for (const [livePath, liveId] of Array.from(liveFolderIdsByPath.entries())) {
          if (livePath.endsWith(`/${expectedPath}`)) {
            id = liveId;
            break;
          }
        }
      }
      return { groupKey: row.groupKey, id };
    }).filter((item): item is { groupKey: string; id: string } => Boolean(item.id));

    if (liveFolders.length > 0) {
      const BATCH = 2;
      for (let i = 0; i < liveFolders.length; i += BATCH) {
        const batch = liveFolders.slice(i, i + BATCH);
        await Promise.all(batch.map(folder => this._refreshFolderFiles(folder.id, folder.groupKey, true)));
      }
    }

    // Also walk the vessel folder tree via Graph delta merge to ensure all files in all subfolders are ingested
    void this._mergeLiveSharePointFiles([vesselName]).catch(() => undefined);

    // Scan the vessel's main folder for out-of-structure subfolders.
    // Resolve the actual SPO path from spoFolderMap instead of guessing.
    const mainFolder = this.state.docMainFolder || 'Technical & Crewing';
    const normVessel = vesselName.trim().toLowerCase();
    const normMain = mainFolder.trim().toLowerCase();

    // Find the live main-folder node for this vessel from the delta map.
    let mainFolderNodeId: string | null = null;
    let resolvedFolderPath: string | null = null;
    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder || node.deleted) continue;
      const segs = (node.serverRelativePath || '').split('/').filter(Boolean).map(s => s.toLowerCase());
      // Match a node whose last segment is the main folder name and whose
      // parent segment is the vessel name (handles any container depth).
      if (segs.length >= 2 &&
          segs[segs.length - 1] === normMain &&
          segs[segs.length - 2].replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '') === normVessel.replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '')) {
        mainFolderNodeId = node.id;
        resolvedFolderPath = node.serverRelativePath;
        break;
      }

    }

    if (!mainFolderNodeId || !resolvedFolderPath) return;
    const signal = new AbortController().signal;

    try {
      // Use the confirmed live path from spoFolderMap — no guessing needed.
      const cleanPath = resolvedFolderPath.replace(/^\/+/, '');
      let children = await this._getGraphChildren(cleanPath, signal);
      let folderPath = cleanPath;
      if (!children || children.length === 0) return;

      const expectedCats: Record<string, string[]> = Object.fromEntries(
        MAIN_FOLDERS.map(main => [main.name, main.perVesselTree.map(folder => folder.name)])
      );
      const expected = expectedCats[mainFolder] || [];

      const newAnomalies: FolderAnomalyItem[] = [];

      for (const item of children) {
        if (!expected.includes(item.name) && item.name !== 'To be Classified') {
          newAnomalies.push({
            id: Date.now() + Math.floor(Math.random() * 10000),
            drive_item_id: item.id,
            name: item.name,
            item_type: item.isFolder ? 'folder' : 'file',
            anomaly_type: 'subfolder_unmatched',
            department: mainFolder,
            vessel_name: vesselName,
            spo_path: `${folderPath.replace(/^\/+/, '')}/${item.name}`,
            resolved: false,
            detected_at: new Date().toISOString(),
          });
        }
      }

      if (newAnomalies.length > 0) {
        this.setState(prev => {
          const existingPaths = new Set(prev.folderAnomalies.map(a => a.spo_path));
          const filtered = newAnomalies.filter(a => !existingPaths.has(a.spo_path));
          if (filtered.length === 0) return null as any;
          return { folderAnomalies: [...prev.folderAnomalies, ...filtered] };
        }, () => this._syncAnomalyAlerts());
      }
    } catch (err) {
      console.warn('[VesselDMS] _loadFilesForVessel warning:', err);
    }
  }


  public _filesLoadedForFolders: Set<string> = new Set();
  public _refreshFolderFilesInFlight: Set<string> = new Set(); // prevents concurrent duplicate fetches
  private _graphChildrenInFlight: Map<string, Promise<any>> = new Map();
  private _graphChildrenCache: Map<string, { value: any; expiresAt: number }> = new Map();
  // "This path doesn't exist" results, separate from _graphChildrenCache
  // (which only ever holds successes — a thrown request never reaches its
  // cache.set). _refreshFolderFiles's path-lookup fallback (below) guesses
  // several candidate folder layouts per vessel and fires a real Graph
  // request for each one, in order, stopping at the first that resolves.
  // Every call that DOESN'T hit the layout on the first guess re-fires every
  // losing guess again from scratch — repeatedly 404ing the same known-bad
  // URLs against Graph every time a folder is refreshed, which is both the
  // noisy 404 console spam and real load against Graph/SharePoint's rate
  // limits. Remembering a 404 here for a few minutes lets repeat calls skip
  // straight past guesses already known to be wrong.
  private _graphChildrenNotFound: Map<string, number> = new Map();
  private static readonly GRAPH_NOT_FOUND_TTL_MS = 3 * 60 * 1000;

  /** Share one Graph children request across overlapping refreshes. */
  private async _getGraphChildrenResponse(url: string): Promise<any> {
    const cached = this._graphChildrenCache.get(url);
    if (cached && cached.expiresAt > Date.now()) return cached.value;
    const notFoundAt = this._graphChildrenNotFound.get(url);
    if (notFoundAt !== undefined) {
      if (Date.now() - notFoundAt < VesselEmail.GRAPH_NOT_FOUND_TTL_MS) {
        const err: any = new Error('itemNotFound (cached)');
        err.statusCode = 404;
        err.code = 'itemNotFound';
        throw err;
      }
      this._graphChildrenNotFound.delete(url);
    }
    const active = this._graphChildrenInFlight.get(url);
    if (active) return active;
    const request = (async () => {
      try {
        const result: any = await this.props.graphClient!.api(url).get();
        this._graphChildrenCache.set(url, { value: result, expiresAt: Date.now() + 5000 });
        return result;
      } catch (err: any) {
        // Only remember genuine "doesn't exist" results — a transient
        // network/5xx error should still be retried next time, not
        // permanently (well, for GRAPH_NOT_FOUND_TTL_MS) treated as 404.
        if (err?.statusCode === 404 || err?.code === 'itemNotFound') {
          this._graphChildrenNotFound.set(url, Date.now());
        }
        throw err;
      }
    })();
    this._graphChildrenInFlight.set(url, request);
    try {
      return await request;
    } finally {
      this._graphChildrenInFlight.delete(url);
    }
  }

  public async _refreshFolderFiles(
    uploadFolderId: string,
    groupKey: string,
    force: boolean = false,
    isConfirmedGraphFolderId: boolean = false,
  ): Promise<void> {
    if (this._isUnmounted) return;
    if (!uploadFolderId) return;
    if (/^f\d+$/.test(uploadFolderId)) return;
    // These IDs exist only in the navigation tree and are not SharePoint
    // drive item IDs. Their contents are loaded by path/vessel-level logic.
    if (/^(kaizen_root|vessels_root|specific_vessels)$/i.test(uploadFolderId)) return;
    if (/^(spo_|live:)/.test(groupKey) || /^01YT4WOQ/.test(uploadFolderId) || /^014ZGIJ/.test(uploadFolderId)) {
      this._filesLoadedForFolders.add(uploadFolderId);
      return;
    }
    if (this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === uploadFolderId.toLowerCase())) {
      this._filesLoadedForFolders.add(uploadFolderId);
      return;
    }
    if (!force && this._filesLoadedForFolders.has(uploadFolderId)) return;
    // Prevent concurrent duplicate fetches for the same folder
    if (this._refreshFolderFilesInFlight.has(uploadFolderId)) return;
    this._refreshFolderFilesInFlight.add(uploadFolderId);
    const refreshToken = ++this._folderRefreshSeq;
    const refreshKeys = new Set<string>([uploadFolderId, groupKey].filter(Boolean));
    refreshKeys.forEach(key => this._latestFolderRefreshTokenByKey.set(key, refreshToken));

    // Keep path-like IDs as-is. Refresh resolves them through live Graph path
    // lookups and spoFolderMap, which avoids backend resolve-only 404 noise
    // for newly created folders during index lag.
    let resolvedId = uploadFolderId;

    // Prefer live Graph API fetch so files uploaded directly in SPO are visible immediately
    const { graphClient, siteId, driveId } = this.props;
    let fileItems: any[] = [];
    let folderItems: any[] = [];
    let effectiveFolderNodeId: string | null = null;
    let fetchedFromGraph = false;

    const isUploadFolderIdLive = this.state.spoFolderMap.has(uploadFolderId);
    const isResolvedIdLiveEarly = this.state.spoFolderMap.has(resolvedId);

    // Prefer unique row key matching first. Backend DB folder IDs can be shared
    // across similarly named branches (e.g. multiple "To be Classified" rows).
    let matchingRow = this.state.rows.find(r => groupKey && r.groupKey === groupKey);
    if (!matchingRow) {
      matchingRow = this.state.rows.find(r =>
        (isUploadFolderIdLive && r.uploadFolderId === uploadFolderId) ||
        (isResolvedIdLiveEarly && r.uploadFolderId === resolvedId)
      );
    }

    // Strategy 4: if still no match and we have groupKey, try to find by breadcrumb
    if (!matchingRow && groupKey) {
      matchingRow = this.state.rows.find(r => r.groupKey === groupKey);
    }

    const sharePointFolderPath = matchingRow
      ? this._sharePointFolderPath(matchingRow.subFolderPath, '')
      : '';
    const cleanP = (p: string) => (p || '').replace(/^[\s>/]+|[\s>/]+$/g, '').replace(/\s*>\s*/g, '/');
    const normalizeRelPath = (p: string) => (p || '').replace(/^\/+|\/+$/g, '').toLowerCase();
    const resolveLiveFolderIdByPath = (p: string): string | null => {
      const target = normalizeRelPath(p);
      if (!target) return null;
      for (const [id, node] of Array.from(this.state.spoFolderMap.entries())) {
        if (!node?.isFolder) continue;
        const live = normalizeRelPath(node.serverRelativePath || '');
        if (live === target || live.endsWith(`/${target}`)) return id;
      }
      return null;
    };
    const directPath = cleanP(uploadFolderId);
    const resolvedPath = cleanP(resolvedId);
    const spPath = cleanP(sharePointFolderPath);
    const pathSegments = spPath ? spPath.split('/').filter(Boolean) : [];

    // This ID was obtained from the active drive's delta response, unlike a
    // database folder ID which can point to another drive.
    const isLiveDriveFolderId = this.state.spoFolderMap.has(uploadFolderId);

    // Also check if resolvedId is in the live drive folder map
    const isResolvedIdLive = this.state.spoFolderMap.has(resolvedId);

    if (graphClient && siteId && driveId) {
      const looksLikeDriveItemId = (s: string) => !s.includes('/') && !s.includes('>') && (/^01[A-Za-z0-9]{20,}/.test(s) || /^[A-Za-z0-9_-]{24,}$/.test(s));
      const validPaths = [spPath, directPath, resolvedPath].filter(p => Boolean(p) && !looksLikeDriveItemId(p));
      if ((!isLiveDriveFolderId && pathSegments.length >= 2) || validPaths.length > 0) {
        const rawBasePaths = Array.from(new Set(validPaths));
        const tryPathsSet = new Set<string>();

        for (const basePath of rawBasePaths) {
          tryPathsSet.add(basePath);

          const segs = basePath.split('/').filter(Boolean);
          if (segs.length >= 2) {
            const isFirstMain = this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === segs[0].toLowerCase());
            const isSecondMain = this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === segs[1].toLowerCase());

            if (isFirstMain && !isSecondMain) {
              const mainF = segs[0];
              const vesselF = segs[1];
              const rest = segs.slice(2).join('/');
              tryPathsSet.add(`${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
              tryPathsSet.add(`Vessels/${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
              tryPathsSet.add(`Vessels/Specific Vessels/${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
              tryPathsSet.add(`${mainF}/Vessels/${vesselF}${rest ? `/${rest}` : ''}`);
            } else if (isSecondMain && !isFirstMain) {
              const vesselF = segs[0];
              const mainF = segs[1];
              const rest = segs.slice(2).join('/');
              tryPathsSet.add(`${mainF}/${vesselF}${rest ? `/${rest}` : ''}`);
              tryPathsSet.add(`Vessels/${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
              tryPathsSet.add(`Vessels/Specific Vessels/${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
              tryPathsSet.add(`${mainF}/Vessels/${vesselF}${rest ? `/${rest}` : ''}`);
            }
          }

          if (segs.length >= 3 && segs[0].toLowerCase() === 'vessels') {
            const hasSpecific = segs[1]?.toLowerCase() === 'specific vessels';
            const vesselIdx = hasSpecific ? 2 : 1;
            const mainIdx = vesselIdx + 1;
            const vesselF = segs[vesselIdx];
            const mainF = segs[mainIdx];
            const rest = segs.slice(mainIdx + 1).join('/');
            if (vesselF && mainF) {
              const isMain = this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === mainF.toLowerCase());
              if (isMain) {
                tryPathsSet.add(`${mainF}/${vesselF}${rest ? `/${rest}` : ''}`);
                tryPathsSet.add(`${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
                tryPathsSet.add(`${mainF}/Vessels/${vesselF}${rest ? `/${rest}` : ''}`);
                tryPathsSet.add(`Vessels/${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
                tryPathsSet.add(`Vessels/Specific Vessels/${vesselF}/${mainF}${rest ? `/${rest}` : ''}`);
              }
            }
          }
        }

        const tryPaths = Array.from(tryPathsSet).map(p => p.replace(/^\/+|\/+$/g, '')).filter(Boolean);
        for (const tryPath of tryPaths) {
          try {
            const encodedPath = tryPath.split('/').filter(Boolean).map(s => encodeURIComponent(s)).join('/');
            const pathUrl = `/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/children?$select=id,name,size,createdDateTime,lastModifiedDateTime,file,folder&$top=200`;
            const result: any = await this._getGraphChildrenResponse(pathUrl);
            const items: any[] = result?.value ?? [];
            const hits = items.filter((i: any) => !i.folder);
            if (hits.length > 0 || (result?.value && Array.isArray(result.value))) {
              fileItems = hits;
              folderItems = items.filter((i: any) => !!i.folder);
              if (!effectiveFolderNodeId) {
                effectiveFolderNodeId =
                  resolveLiveFolderIdByPath(tryPath) ||
                  resolveLiveFolderIdByPath(spPath) ||
                  resolveLiveFolderIdByPath(resolvedPath) ||
                  resolveLiveFolderIdByPath(directPath);
              }
              fetchedFromGraph = true;
              console.log(`[VesselDMS] _refreshFolderFiles: fetched ${fileItems.length} files and ${folderItems.length} folders via path lookup for "${tryPath}"`);
              break;
            }
          } catch (pathLookupError) {
            // try next path
          }
        }
      }

      // Use an item-ID lookup when the ID is confirmed to be from the active live drive.
      // isConfirmedGraphFolderId=true means the caller (upload handler) verified this is a real Graph drive item ID.
      const liveIdFromBreadcrumb = matchingRow ? this._getLiveSharePointFolderId(matchingRow.subFolderPath) : null;
      const liveIdFromUploadPath = uploadFolderId.includes('>') ? this._getLiveSharePointFolderId(uploadFolderId) : null;
      let targetGraphId: string | null = null;
      if (isConfirmedGraphFolderId) {
        // Upload succeeded via Graph - uploadFolderId is the real drive item ID
        targetGraphId = uploadFolderId;
      } else if (isLiveDriveFolderId) {
        // uploadFolderId was found in spoFolderMap (from delta sync)
        targetGraphId = uploadFolderId;
      } else if (isResolvedIdLive) {
        // resolvedId was found in spoFolderMap
        targetGraphId = resolvedId;
      } else {
        // Fall back to breadcrumb resolution
        targetGraphId = liveIdFromBreadcrumb || liveIdFromUploadPath;
      }
      if (!fetchedFromGraph && targetGraphId && !targetGraphId.includes('/')) {
        try {
          const url = `/sites/${siteId}/drives/${driveId}/items/${targetGraphId}/children?$select=id,name,size,createdDateTime,lastModifiedDateTime,file,folder&$top=200`;
          const result: any = await this._getGraphChildrenResponse(url);
          const items: any[] = result?.value ?? [];
          folderItems = items.filter((i: any) => !!i.folder);
          fileItems = items.filter((i: any) => !i.folder); // only files, not subfolders
          effectiveFolderNodeId = targetGraphId;
          fetchedFromGraph = true;
          console.log(`[VesselDMS] _refreshFolderFiles: fetched ${fileItems.length} files via ID lookup for "${targetGraphId}"`);
        } catch (idLookupError) {
          console.warn(`[VesselDMS] failed to load confirmed live folder ID "${targetGraphId}"`, idLookupError);
        }
      }
    }

    if (!fetchedFromGraph) {
      // Path-like IDs are not valid for /api/folders/{id}/children and will
      // cause avoidable 404 loops. We already attempted Graph path lookups.
      if (resolvedId.includes('/') || resolvedId.includes('>')) {
        console.warn(`[VesselDMS] _refreshFolderFiles: no Graph result for path "${resolvedId}"; skipping REST children fallback`);
        this._refreshFolderFilesInFlight.delete(uploadFolderId);
        return;
      }
      // Skip REST fallback for numeric-only IDs — these are DB row IDs, not SharePoint drive item IDs.
      // Passing them to /api/folders/{id}/children causes Graph 404s and a polling loop.
      if (/^\d+$/.test(resolvedId)) {
        console.warn(`[VesselDMS] _refreshFolderFiles: skipping REST fallback for numeric DB ID "${resolvedId}" — not a valid drive item ID`);
        this._refreshFolderFilesInFlight.delete(uploadFolderId);
        return;
      }
      try {
        const children: any[] = await this._fetchJson(`${this._base()}/api/folders/${encodeURIComponent(resolvedId)}/children`);
        fileItems = (children || []).filter((c: any) => c.kind === 'file');
        console.log(`[VesselDMS] _refreshFolderFiles: fetched ${fileItems.length} files via REST API for "${resolvedId}"`);
      } catch (err) {
        console.warn(`[VesselDMS] failed to load children for folder "${resolvedId}"`, err);
        this._refreshFolderFilesInFlight.delete(uploadFolderId);
        return;
      }
    }

    // Mark folder as loaded to prevent component re-render loops on empty folders
    this._filesLoadedForFolders.add(uploadFolderId);
    if (resolvedId && resolvedId !== uploadFolderId) this._filesLoadedForFolders.add(resolvedId);
    this._refreshFolderFilesInFlight.delete(uploadFolderId);
    if (resolvedId) this._refreshFolderFilesInFlight.delete(resolvedId);

    const parsedUploads = fileItems.map((f: any) => ({
      name: f.name || f.displayName,
      size: f.size ? `${(f.size / 1024).toFixed(1)} KB` : '—',
      date: f.createdDateTime ? new Date(f.createdDateTime).toLocaleString([], { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : (f.lastModifiedDateTime ? new Date(f.lastModifiedDateTime).toLocaleString([], { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : (f.modified || 'Today')),
      pending: false,
      id: f.id,
      uploadedAt: f.createdDateTime ? Date.parse(f.createdDateTime) : (f.lastModifiedDateTime ? Date.parse(f.lastModifiedDateTime) : (f.modified ? Date.parse(f.modified) : undefined)),
    }));

    const staleBeforeSetState = Array.from(refreshKeys).some(key => this._latestFolderRefreshTokenByKey.get(key) !== refreshToken);
    if (staleBeforeSetState) {
      console.log('[VesselDMS] _refreshFolderFiles: stale refresh result ignored before state merge', { uploadFolderId, groupKey, resolvedId });
      return;
    }

    if (!effectiveFolderNodeId) {
      effectiveFolderNodeId =
        resolveLiveFolderIdByPath(spPath) ||
        resolveLiveFolderIdByPath(resolvedPath) ||
        resolveLiveFolderIdByPath(directPath) ||
        (matchingRow?.subFolderPath ? this._getLiveSharePointFolderId(matchingRow.subFolderPath) : null);
    }

    this.setState(prev => {
      const staleDuringSetState = Array.from(refreshKeys).some(key => this._latestFolderRefreshTokenByKey.get(key) !== refreshToken);
      if (staleDuringSetState) return null;
      // Merge fetched parsedUploads with any existing local uploads so we don't wipe out freshly uploaded files.
      // Also check groupKey so a just-uploaded file stored there isn't lost due to React state batching.
      const matchingRow = prev.rows.find(r =>
        (groupKey && r.groupKey === groupKey) ||
        (isUploadFolderIdGraph && r.uploadFolderId === uploadFolderId) ||
        (isResolvedIdGraph && r.uploadFolderId === resolvedId)
      );

      // Collect all locally-cached uploads across every key that could hold this folder's files.
      // This is critical for the post-upload consistency window where Graph returns empty.
      // IMPORTANT: Do NOT read from numeric backend DB IDs (uploadFolderId/resolvedId) as they may be shared
      // across multiple "To be Classified" rows from different categories. Only read from unique keys:
      // - groupKey (unique per row: vessel||group||category||subCategory||breadcrumb)
      // - normSub (normalized subFolderPath)
      // - dedupeKey (normalized subFolderPath or groupKey)
      // - resolvedId ONLY if it's a real Graph drive item ID (in spoFolderMap)
      const isUploadFolderIdGraph = this.state.spoFolderMap.has(uploadFolderId);
      const isResolvedIdGraph = this.state.spoFolderMap.has(resolvedId);

      const allCachedUploads = [
        // Read from groupKey (unique per row)
        ...(matchingRow ? (prev.uploadedFilesByFolder[matchingRow.groupKey] || []) : []),
        ...(groupKey ? (prev.uploadedFilesByFolder[groupKey] || []) : []),
        // Read from normalized subFolderPath (unique per folder path)
        // Only read from uploadFolderId/resolvedId if they are REAL Graph drive item IDs
        ...(isUploadFolderIdGraph ? (prev.uploadedFilesByFolder[uploadFolderId] || []) : []),
        ...(isResolvedIdGraph ? (prev.uploadedFilesByFolder[resolvedId] || []) : []),
      ];
      const cachedByName = new Map<string, any>();
      allCachedUploads.forEach((file: any) => {
        if (file?.name && !cachedByName.has(file.name.toLowerCase())) {
          cachedByName.set(file.name.toLowerCase(), file);
        }
      });
      const mergedUploads: Array<{ name: string; size: string; date: string; pending: boolean; id: string; uploadedAt?: number }> = parsedUploads.map(file => ({
        ...file,
        uploadedAt: cachedByName.get(file.name.toLowerCase())?.uploadedAt,
      }));
      for (const ex of allCachedUploads) {
        if (!mergedUploads.some(u => u.name === ex.name)) {
          mergedUploads.push({
            name: ex.name,
            size: ex.size || '—',
            date: ex.date || 'Today',
            pending: Boolean(ex.pending),
            id: (ex as any).id || ex.name,
            uploadedAt: (ex as any).uploadedAt,
          });
        }
      }

      const updatedByFolder: Record<string, any[]> = {
        ...prev.uploadedFilesByFolder,
        [groupKey]: mergedUploads,
      };
      if (isUploadFolderIdGraph && uploadFolderId) {
        updatedByFolder[uploadFolderId] = mergedUploads;
        updatedByFolder[uploadFolderId.toLowerCase()] = mergedUploads;
      }
      if (isResolvedIdGraph && resolvedId && resolvedId !== uploadFolderId) {
        updatedByFolder[resolvedId] = mergedUploads;
        updatedByFolder[resolvedId.toLowerCase()] = mergedUploads;
      }
      if (directPath && directPath.includes('/')) {
        updatedByFolder[directPath] = mergedUploads;
        updatedByFolder[directPath.toLowerCase()] = mergedUploads;
      }
      if (matchingRow) {
        if (matchingRow.subFolderPath) {
          updatedByFolder[matchingRow.subFolderPath] = mergedUploads;
          updatedByFolder[matchingRow.subFolderPath.trim().toLowerCase()] = mergedUploads;
        }
      }

      const existingBaseRows = prev.rows.filter(r =>
        // Only match by groupKey or real Graph drive item IDs
        r.groupKey === groupKey ||
        (isUploadFolderIdGraph && r.uploadFolderId === uploadFolderId) ||
        (isResolvedIdGraph && r.uploadFolderId === resolvedId)
      );

      // Also store under the base row's groupKey if different
      if (existingBaseRows.length > 0) {
        const baseRow0 = existingBaseRows[0];
        if (baseRow0.groupKey && baseRow0.groupKey !== groupKey) updatedByFolder[baseRow0.groupKey] = mergedUploads;
        if (baseRow0.subFolderPath) {
          updatedByFolder[baseRow0.subFolderPath] = mergedUploads;
          updatedByFolder[baseRow0.subFolderPath.trim().toLowerCase()] = mergedUploads;
        }
        // Do NOT write to baseRow0.uploadFolderId if it's a backend DB ID
        if (baseRow0.uploadFolderId && this.state.spoFolderMap.has(baseRow0.uploadFolderId)) {
          updatedByFolder[baseRow0.uploadFolderId] = mergedUploads;
        }
      }

      let newRows: FlatRow[] = prev.rows;
      let nextSpoFolderMap = prev.spoFolderMap;

      // Keep the folder tree live: merge folder + file children from Graph into
      // the current folder node so Folder View reflects latest SPO changes.
      const candidateLiveFolderId = effectiveFolderNodeId ||
        (this.state.spoFolderMap.has(uploadFolderId) ? uploadFolderId : null) ||
        (this.state.spoFolderMap.has(resolvedId) ? resolvedId : null) ||
        null;
      if (candidateLiveFolderId && prev.spoFolderMap.has(candidateLiveFolderId) && (folderItems.length > 0 || fileItems.length > 0)) {
        const existingNode = prev.spoFolderMap.get(candidateLiveFolderId)!;
        const mergedChildren = new Map<string, any>();
        (existingNode.children || []).forEach(ch => mergedChildren.set(ch.id, ch));

        for (const it of [...folderItems, ...fileItems]) {
          const id = it?.id;
          const name = it?.name;
          if (!id || !name) continue;
          const isFolderNode = !!it.folder;
          mergedChildren.set(id, {
            id,
            name,
            parentId: candidateLiveFolderId,
            isFolder: isFolderNode,
            serverRelativePath: `${(existingNode.serverRelativePath || '').replace(/\/+$/, '')}/${name}`,
            children: isFolderNode ? (mergedChildren.get(id)?.children || []) : [],
            size: typeof it.size === 'number' ? it.size : undefined,
            createdDateTime: it.createdDateTime,
            lastModifiedDateTime: it.lastModifiedDateTime,
            deleted: false,
          });
        }

        const updatedNode = {
          ...existingNode,
          children: Array.from(mergedChildren.values()),
        };
        nextSpoFolderMap = new Map(prev.spoFolderMap);
        nextSpoFolderMap.set(candidateLiveFolderId, updatedNode);
      }

      if (existingBaseRows.length > 0) {
        const baseRow = existingBaseRows[0];
        const isUploadFolderIdGraph = this.state.spoFolderMap.has(uploadFolderId);
        const isResolvedIdGraph = this.state.spoFolderMap.has(resolvedId);
        const existingGroupUploads = [
          ...(prev.uploadedFilesByFolder[baseRow.groupKey] || []),
          // Only read from uploadFolderId/resolvedId if they are real Graph drive item IDs
          ...(isUploadFolderIdGraph ? (prev.uploadedFilesByFolder[uploadFolderId] || []) : []),
          ...(isResolvedIdGraph ? (prev.uploadedFilesByFolder[resolvedId] || []) : []),
          ...(groupKey && groupKey !== baseRow.groupKey ? (prev.uploadedFilesByFolder[groupKey] || []) : []),
        ];
        const mergedGroupUploads = [...mergedUploads];
        for (const ex of existingGroupUploads) {
          if (!mergedGroupUploads.some(u => u.name === ex.name)) {
            mergedGroupUploads.push({
              name: ex.name,
              size: ex.size || '—',
              date: ex.date || 'Today',
              pending: Boolean(ex.pending),
              id: (ex as any).id || ex.name,
            });
          }
        }
        updatedByFolder[baseRow.groupKey] = mergedGroupUploads;

        // Keep all rows except the base row (which is re-added via mappedRows below).
        // Filter only by groupKey (unique per row) — never by shared backend DB IDs.
        const otherFolderRows = prev.rows.filter(r => r.groupKey !== baseRow.groupKey);

        // CRITICAL FIX: If fileItems is empty but we have local/existing uploads (including pending),
        // NEVER overwrite rows - Graph may not be consistent yet.
        // The mergedGroupUploads includes local cache, so if it has content, use it.
        // If mergedGroupUploads is also empty, preserve the existing rows instead of clearing them.
        const activeFiles = mergedGroupUploads.length > 0 ? mergedGroupUploads : parsedUploads;

        console.log('[VesselDMS] _refreshFolderFiles merge:', {
          uploadFolderId,
          resolvedId,
          groupKey,
          fileItemsCount: fileItems.length,
          parsedUploadsCount: parsedUploads.length,
          mergedGroupUploadsCount: mergedGroupUploads.length,
          activeFilesCount: activeFiles.length,
          existingBaseRowsCount: existingBaseRows.length,
          baseRowGroupKey: baseRow?.groupKey,
          uploadedFilesByFolderKeys: Object.keys(prev.uploadedFilesByFolder),
        });

        if (activeFiles.length === 0) {
          // Graph may not be consistent yet — never overwrite rows when we have nothing to show.
          // Also do NOT overwrite uploadedFilesByFolder keys that already have files — a concurrent
          // successful refresh may have just written them and this stale empty result must not clobber them.
          console.log(`[VesselDMS] _refreshFolderFiles: Graph returned empty for folder "${uploadFolderId}", preserving existing rows`);
          const safeByFolder = { ...prev.uploadedFilesByFolder };
          for (const [k, v] of Object.entries(updatedByFolder)) {
            if (!safeByFolder[k] || safeByFolder[k].length === 0) {
              safeByFolder[k] = v;
            }
          }
          return { rows: prev.rows, uploadedFilesByFolder: safeByFolder, spoFolderMap: nextSpoFolderMap };
        } else {
          const mappedRows = activeFiles.map((f: any, i: number) => ({
            ...baseRow,
            srNo: i === 0 ? baseRow.srNo : `${baseRow.srNo}.${i + 1}`,
            fileName: f.name || f.displayName || null,
            fileId: f.id || null,
            filePending: Boolean(f.pending),
          }));
          newRows = [...otherFolderRows, ...mappedRows];
        }
      }
      return { rows: newRows, uploadedFilesByFolder: updatedByFolder, spoFolderMap: nextSpoFolderMap };
    });
  }

  /** Open the file-delete confirmation dialog for one or more files in a folder. */
  public _openFileDeleteDialog(
    files: Array<{ id: string; name: string; folderId: string; folderPath: string; vesselName?: string; siteId?: string; driveId?: string }>
  ): void {
    const validFiles = files.filter(f => f.id && f.name);
    if (validFiles.length === 0) {
      console.warn('[VesselDMS] _openFileDeleteDialog: no valid files to delete', files);
      return;
    }
    console.log('[VesselDMS] _openFileDeleteDialog opening for', validFiles.map(f => f.name));
    this.setState({
      fileDeleteDialog: {
        files: validFiles,
        selected: new Set(validFiles.map(f => f.id)),
        busy: false,
        error: null,
      },
    });
  }

  /** Soft-delete selected files via Graph (moves to SharePoint Recycle Bin) and update state. */
  public async _deleteSelectedFiles(): Promise<void> {
    const dialog = this.state.fileDeleteDialog;
    if (!dialog || dialog.selected.size === 0) return;
    const { graphClient, siteId, driveId } = this.props;

    this.setState(prev => ({
      fileDeleteDialog: prev.fileDeleteDialog ? { ...prev.fileDeleteDialog, busy: true, error: null } : null,
    }));

    // Re-read from state after setState to get the latest snapshot
    const { files, selected } = dialog;
    const toDelete = files.filter(f => selected.has(f.id));
    if (toDelete.length === 0) {
      this.setState(prev => ({
        fileDeleteDialog: prev.fileDeleteDialog ? { ...prev.fileDeleteDialog, busy: false } : null,
      }));
      return;
    }

    const errors: string[] = [];
    const deletedIds: string[] = [];

    for (const file of toDelete) {
      try {
        // Files browsed in the SharePoint Sites view live in that site's own
        // drive, not necessarily the web part's: use the file's own site/drive
        // when the caller supplied them.
        const fileSiteId = file.siteId || siteId;
        const fileDriveId = file.driveId || driveId;
        // Graph DELETE on a drive item moves it to SharePoint's own Recycle Bin (soft delete).
        // Only use Graph for real drive item IDs — numeric-only IDs are backend DB IDs.
        if (graphClient && fileSiteId && fileDriveId && file.id && !/^file_/.test(file.id) && !/^\d+$/.test(file.id)) {
          // DELETE /drives/{driveId}/items/{itemId} → soft delete (moves to SPO Recycle Bin)
          await graphClient.api(`/sites/${fileSiteId}/drives/${fileDriveId}/items/${file.id}`).delete();
        } else {
          // Fallback: backend soft-delete
          const res = await fetch(`${this._base()}/api/files/${encodeURIComponent(file.id)}`, {
            method: 'DELETE', headers: this._headers(),
          });
          if (!res.ok && res.status !== 202) throw new Error(`HTTP ${res.status}`);
        }
        deletedIds.push(file.id);

        const usedGraph = graphClient && fileSiteId && fileDriveId && file.id && !/^file_/.test(file.id) && !/^\d+$/.test(file.id);
        let recycleBinItemId: string | undefined;
        if (usedGraph && this.props.siteUrl) {
          try {
            // Graph /recycleBin is not supported in v1.0 - use SharePoint REST API instead.
            const nameLower = file.name.toLowerCase();
            const rbRes = await fetch(
              `${this.props.siteUrl}/_api/site/RecycleBin?$filter=LeafName eq '${encodeURIComponent(file.name)}'&$select=Id,LeafName&$top=10`,
              { headers: { Accept: 'application/json;odata=nometadata' } }
            );
            if (rbRes.ok) {
              const rbData = await rbRes.json();
              const match = (rbData?.value ?? []).find((r: any) => (r.LeafName || '').toLowerCase() === nameLower);
              if (match?.Id) recycleBinItemId = match.Id;
            }
          } catch { /* non-critical */ }
        }
                const activeSiteObj = (this.state.documentSites || []).find(s => s.site_key === this.state.activeDocumentSite);
        const fileVessel = this.state.vessels.find(v => (v.name || '').toLowerCase() === (file.vesselName || '').toLowerCase());
        const siteNames = (fileVessel?.provisioned_site_ids || []).map(sk => {
          const matched = (this.state.documentSites || []).find(s => s.site_key === sk);
          return matched?.sp_site_name || sk;
        }).filter(Boolean);
        const resolvedSiteName = activeSiteObj?.sp_site_name || (siteNames.length > 0 ? siteNames.join(', ') : (this.props.siteUrl || 'SharePoint'));

        const deletedNode: import('./types/ui').DeletedNode = {
          id: file.id,
          name: file.name,
          kind: 'file',
          item_type: 'file',
          original_path: file.folderPath,
          deleted_at: new Date().toISOString(),
          ext: file.name.split('.').pop(),
          in_spo_recycle_bin: Boolean(usedGraph),
          recycle_bin_item_id: recycleBinItemId,
          site_name: resolvedSiteName,
          site_key: activeSiteObj?.site_key || this.state.activeDocumentSite || undefined,
        };
        this._logDeletion({
          item_type: 'file',
          drive_item_id: file.id,
          name: file.name,
          original_path: file.folderPath,
          site_name: resolvedSiteName,
          site_key: activeSiteObj?.site_key || this.state.activeDocumentSite || null,
          reason: dialog.reason || null,
        });
        if (usedGraph) this._appDeletedItemIds.add(file.id);
        this._handleSpoDocumentDeletion({
          id: file.id,
          name: file.name,
          parentId: null,
          isFolder: false,
          serverRelativePath: file.folderPath,
          children: [],
        }, {
          // Deleted from inside the app by the current session — attribution
          // is already known, no need to wait on the backend lookup.
          deletedByEmail: this.props.userEmail || null,
          deletedByName: this.props.userDisplayName || null,
          siteName: resolvedSiteName,
          originalPath: file.folderPath,
        });
        this.setState(prev => ({ recycleBin: [...prev.recycleBin, deletedNode] }));
      } catch (err: any) {
        errors.push(`${file.name}: ${err?.message || 'Delete failed'}`);
      }
    }

    if (errors.length > 0) {
      this.setState(prev => ({
        fileDeleteDialog: prev.fileDeleteDialog
          ? { ...prev.fileDeleteDialog, busy: false, error: errors.join(' | ') }
          : null,
      }));
      return;
    }

    // Drop the deleted files from the cached SharePoint folder listings so the
    // Sites / Folder view updates immediately (matched by item id).
    const deletedIdSet = new Set(deletedIds.map(String));
    this._siteFolderItemsCache.forEach((entry, key) => {
      const kept = (entry.items as any[]).filter(it => !deletedIdSet.has(String(it?.id ?? '')));
      if (kept.length !== entry.items.length) this._siteFolderItemsCache.set(key, { ...entry, items: kept } as any);
    });

    // Remove deleted files from uploadedFilesByFolder and rows
    const deletedNames = new Set(toDelete.map(f => f.name.toLowerCase()));
    this.setState(prev => {
      const updatedByFolder: Record<string, any[]> = {};
      for (const [key, files] of Object.entries(prev.uploadedFilesByFolder)) {
        updatedByFolder[key] = (files as any[]).filter((f: any) => !deletedNames.has((f.name || '').toLowerCase()));
      }
      const updatedRows = prev.rows.map(r =>
        r.fileName && deletedNames.has(r.fileName.toLowerCase())
          ? { ...r, fileName: null, fileId: null }
          : r
      );
      return { rows: updatedRows, uploadedFilesByFolder: updatedByFolder, fileDeleteDialog: null };
    });

    void this._syncScheduler?.triggerNow().catch(() => undefined);
  }

  public _renderFileDeleteDialog(): React.ReactElement | null {
    const { fileDeleteDialog } = this.state;
    if (!fileDeleteDialog) return null;
    const { files, selected, busy, error } = fileDeleteDialog;

    const toggle = (id: string): void => {
      this.setState(prev => {
        if (!prev.fileDeleteDialog) return null as any;
        const next = new Set(prev.fileDeleteDialog.selected);
        if (next.has(id)) next.delete(id); else next.add(id);
        return { fileDeleteDialog: { ...prev.fileDeleteDialog, selected: next } };
      });
    };

    // Collect additional files from the same folder(s) not yet in the dialog
    const dialogFolderIds = new Set(files.map(f => f.folderId).filter(Boolean));
    const dialogFileIds = new Set(files.map(f => f.id));
    const additionalFiles: Array<{ id: string; name: string; folderId: string; folderPath: string }> = [];
    for (const [key, folderFiles] of Object.entries(this.state.uploadedFilesByFolder)) {
      if (!dialogFolderIds.has(key)) continue;
      for (const f of (folderFiles as any[])) {
        if (f?.id && f?.name && !dialogFileIds.has(f.id)) {
          additionalFiles.push({ id: f.id, name: f.name, folderId: key, folderPath: files[0]?.folderPath || '' });
          dialogFileIds.add(f.id);
        }
      }
    }

    const addFileToDialog = (f: { id: string; name: string; folderId: string; folderPath: string }): void => {
      this.setState(prev => {
        if (!prev.fileDeleteDialog) return null as any;
        return { fileDeleteDialog: { ...prev.fileDeleteDialog, files: [...prev.fileDeleteDialog.files, f] } };
      });
    };

    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

    return (
      <div style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 100001,
        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 20,
      }}>
        <div style={{
          // Fallbacks (2nd arg to each var()) keep this popup fully opaque
          // and legible even if it ever renders before/outside the
          // [data-vessel-theme] scope that defines --vdms-* (see render()'s
          // body-attribute sync above for the actual root-cause fix) —
          // previously an unresolved var() left `background`/`color`
          // unset, so the dialog blended into the dark backdrop overlay.
          background: 'var(--vdms-surface, #ffffff)', borderRadius: 16, padding: isMobile ? '16px 14px' : '28px 32px', width: isMobile ? '95vw' : 'auto', minWidth: isMobile ? 0 : 420, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto',
          boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 20 }} /></div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--vdms-text, #0f172a)' }}>Delete Files</div>
              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted, #64748b)', marginTop: 2 }}>Check files to move to Recycle Bin</div>
            </div>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 12, maxHeight: 260, overflowY: 'auto' }}>
            {files.map(f => (
              <label key={f.id} style={{
                display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px',
                borderRadius: 8, border: `1.5px solid ${selected.has(f.id) ? '#ef4444' : 'var(--vdms-border, #e2e8f0)'}`,
                background: selected.has(f.id) ? '#fff5f5' : 'var(--vdms-surface-alt, #f8fafc)', cursor: 'pointer',
              }}>
                <input
                  type="checkbox"
                  checked={selected.has(f.id)}
                  onChange={() => toggle(f.id)}
                  style={{ width: 16, height: 16, accentColor: '#ef4444', cursor: 'pointer' }}
                />
                <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 16 }} />
                <span style={{ fontSize: 13, fontWeight: 600, color: 'var(--vdms-text, #0f172a)', flex: 1, wordBreak: 'break-all' }}>{f.name}</span>
              </label>
            ))}
          </div>

          {/* Additional files from the same folder that can be added */}
          {additionalFiles.length > 0 && (
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: 'var(--vdms-text-muted, #64748b)', textTransform: 'uppercase', marginBottom: 6 }}>
                + Add more files from this folder
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 140, overflowY: 'auto' }}>
                {additionalFiles.map(f => (
                  <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', borderRadius: 8, border: '1px dashed var(--vdms-border, #e2e8f0)', background: 'var(--vdms-surface-alt, #f8fafc)' }}>
                    <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 14 }} />
                    <span style={{ fontSize: 12, color: 'var(--vdms-text, #0f172a)', flex: 1, wordBreak: 'break-all' }}>{f.name}</span>
                    <button
                      onClick={() => addFileToDialog(f)}
                      style={{ border: '1px solid var(--vdms-border, #e2e8f0)', background: 'var(--vdms-surface, #ffffff)', borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: 'var(--vdms-text, #0f172a)', whiteSpace: 'nowrap' }}
                    >
                      + Add
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: 'var(--vdms-text-muted, #64748b)', textTransform: 'uppercase', marginBottom: 6 }}>
            Reason for deletion (optional)
          </label>
          <textarea
            value={fileDeleteDialog.reason || ''}
            onChange={e => this.setState(prev => (
              prev.fileDeleteDialog ? { fileDeleteDialog: { ...prev.fileDeleteDialog, reason: e.target.value } } : null as any
            ))}
            disabled={busy}
            placeholder="Optional — why is this being deleted?"
            rows={2}
            style={{
              width: '100%', boxSizing: 'border-box', borderRadius: 8, border: '1px solid var(--vdms-border, #e2e8f0)',
              padding: '8px 10px', fontSize: 13, fontFamily: 'inherit', resize: 'vertical', marginBottom: 12,
              background: 'var(--vdms-surface-alt, #f8fafc)', color: 'var(--vdms-text, #0f172a)',
            }}
          />

          {error && (
            <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#dc2626', marginBottom: 12 }}>
              <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {error}
            </div>
          )}

          <div style={{ background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#c2410c', marginBottom: 20 }}>
            <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {selected.size} file{selected.size !== 1 ? 's' : ''} will be moved to the Recycle Bin. This can be undone from the Recycle Bin page.
          </div>

          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', flexDirection: isMobile ? 'column' : 'row' }}>
            <button
              onClick={() => this.setState({ fileDeleteDialog: null })}
              disabled={busy}
              style={{ minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: '1px solid var(--vdms-border, #e2e8f0)', background: 'var(--vdms-surface, #ffffff)', fontSize: 13, fontWeight: 600, cursor: 'pointer', color: 'var(--vdms-text, #0f172a)' }}
            >
              Cancel
            </button>
            <button
              onClick={() => void this._deleteSelectedFiles()}
              disabled={busy || selected.size === 0}
              style={{
                minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 600, cursor: busy || selected.size === 0 ? 'not-allowed' : 'pointer',
                background: busy || selected.size === 0 ? '#fca5a5' : '#ef4444', color: '#fff',
              }}
            >
              {busy ? <><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 13 }} /> Deleting…</> : <><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 13 }} /> {`Delete ${selected.size} file${selected.size !== 1 ? 's' : ''}`}</>}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ── Export vessels report (Phase 4): full vessel + folder summary as
  // .xlsx, independent of any Documents/List View filters — see
  // GET /api/reports/vessels-excel.
  public _exportVesselsExcel = async (): Promise<void> => {
    if (this.state.vesselsExcelExportBusy) return;
    this.setState({ vesselsExcelExportBusy: true });
    try {
      const res = await fetch(`${this._base()}/api/reports/vessels-excel`, {
        method: 'GET',
        headers: this._headers(),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data?.detail || data?.message || `Export failed (HTTP ${res.status})`);
      }
      const blob = await res.blob();
      const disposition = res.headers.get('Content-Disposition') || '';
      const match = /filename="?([^";]+)"?/.exec(disposition);
      const filename = match?.[1] || `vessel-dms-report-${new Date().toISOString().slice(0, 10)}.xlsx`;
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (e: any) {
      alert(e?.message || 'Could not export the vessels report. Check your connection and try again.');
    } finally {
      this.setState({ vesselsExcelExportBusy: false });
    }
  };

  // ── Add Folder (Phase 3): create a named subfolder under any folder that
  // already belongs to a vessel (flat root or any depth under it) or is
  // month-driven. Backed by POST /api/folders/{folder_id}/subfolder —
  // see RealBackend.create_subfolder / _execute_create_subfolder. Creates
  // exactly the one folder the user names; no template is applied.
  public _openAddFolderDialog = (options: {
    folderId: string;
    folderLabel: string;
    vesselName: string;
  }): void => {
    const { folderId, folderLabel, vesselName } = options;
    // create_subfolder needs a real SharePoint drive-item id, not a
    // synthetic/breadcrumb key (those contain '/' or aren't resolvable yet).
    if (!folderId || folderId.includes('/')) return;
    this.setState({
      addFolderDialog: { folderId, folderLabel, vesselName, name: '', busy: false, error: null, asVessel: false },
    });
  };

  public _submitAddFolder = async (): Promise<void> => {
    const dialog = this.state.addFolderDialog;
    if (!dialog) return;
    const name = dialog.name.trim();
    if (!name) {
      this.setState({ addFolderDialog: { ...dialog, error: 'Enter a folder name.' } });
      return;
    }
    if (dialog.asVessel) {
      // Vessel creation has its own validated flow (IMO, site/parent-folder
      // picker, etc.) — reopen it instead of re-implementing it here, just
      // pre-filling the name the user already typed.
      this.setState({ addFolderDialog: null });
      this._openCreate(name);
      return;
    }
    this.setState({ addFolderDialog: { ...dialog, busy: true, error: null } });
    try {
      const res = await fetch(`${this._base()}/api/folders/${encodeURIComponent(dialog.folderId)}/subfolder`, {
        method: 'POST',
        headers: this._headers(),
        body: JSON.stringify({ name, user_email: this.props.userEmail || undefined }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok && res.status !== 202) {
        const msg: string = data?.detail || data?.message || `Error ${res.status}`;
        this.setState(prev => ({
          addFolderDialog: prev.addFolderDialog ? { ...prev.addFolderDialog, busy: false, error: msg } : null,
        }));
        return;
      }
      this.setState({ addFolderDialog: null });
      void this._loadData(true);
    } catch (e: any) {
      this.setState(prev => ({
        addFolderDialog: prev.addFolderDialog
          ? { ...prev.addFolderDialog, busy: false, error: e?.message || 'Could not create the folder. Check your connection and try again.' }
          : null,
      }));
    }
  };

  public _renderAddFolderDialog(): React.ReactElement | null {
    const { addFolderDialog } = this.state;
    if (!addFolderDialog) return null;
    const { folderLabel, vesselName, name, busy, error } = addFolderDialog;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

    return (
      <div style={{
        position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.6)', zIndex: 100001,
        backdropFilter: 'blur(4px)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 20,
      }}>
        <div style={{
          background: clay.surface, borderRadius: clay.radiusCard, width: isMobile ? '95vw' : 440, maxWidth: '95vw', overflow: 'hidden',
          boxShadow: '0 20px 50px rgba(15,45,45,0.28), ' + clay.shadowRaised, border: 'none',
        }}>
          <div style={{ padding: '20px 24px 16px', background: clay.accentSoft }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <div style={{ width: 46, height: 46, borderRadius: clay.radiusIcon, background: clay.iconBgGradient, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22, boxShadow: clay.shadowIcon }}>
                <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 22 }} />
              </div>
              <div>
                <div style={{ fontSize: 17, fontWeight: 700, color: clay.text }}>Add Folder</div>
                <div style={{ fontSize: 12, color: clay.accentDark, marginTop: 2, wordBreak: 'break-word' }}>
                  Inside {vesselName ? <strong>{vesselName}</strong> : 'this folder'}{folderLabel ? ` — ${folderLabel}` : ''}
                </div>
              </div>
            </div>
          </div>

          <div style={{ padding: '20px 24px' }}>
            {error && (
              <div style={{ background: '#fde8e0', color: '#9a3f1f', padding: '10px 14px', borderRadius: 12, fontSize: 12, marginBottom: 14 }}>
                <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {error}
              </div>
            )}
            <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: clay.text, marginBottom: 6 }}>
              Folder name
            </label>
            <input
              autoFocus
              value={name}
              disabled={busy}
              onChange={e => this.setState(prev => ({
                addFolderDialog: prev.addFolderDialog ? { ...prev.addFolderDialog, name: e.target.value, error: null } : null,
              }))}
              onKeyDown={e => { if (e.key === 'Enter' && !busy) void this._submitAddFolder(); }}
              placeholder="e.g. Insurance Documents"
              style={{
                width: '100%', padding: '10px 14px', borderRadius: 14, border: 'none', fontSize: 14, outline: 'none',
                boxSizing: 'border-box', color: clay.text, background: clay.bg,
                boxShadow: 'inset 3px 3px 8px rgba(120,190,185,0.35), inset -3px -3px 6px rgba(255,255,255,0.7)',
              }}
            />
            <div style={{ fontSize: 11, color: clay.textMuted, marginTop: 8 }}>
              Only this one folder is created — no other folders are added automatically.
            </div>

            <label style={{
              display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', borderRadius: 12, marginTop: 14,
              background: addFolderDialog.asVessel ? '#eff6ff' : clay.bg, cursor: busy ? 'default' : 'pointer',
              boxShadow: clay.shadowRaised,
            }}>
              <input
                type="checkbox"
                checked={addFolderDialog.asVessel}
                disabled={busy}
                onChange={e => this.setState(prev => ({
                  addFolderDialog: prev.addFolderDialog ? { ...prev.addFolderDialog, asVessel: e.target.checked } : null,
                }))}
                style={{ width: 15, height: 15, accentColor: '#2563eb', cursor: 'pointer' }}
              />
              <span style={{ fontSize: 12.5, color: clay.text }}>
                <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /> This is a new vessel — open the Add Vessel form instead
              </span>
            </label>
          </div>

          <div style={{ padding: '14px 24px 20px', background: clay.bg, display: 'flex', justifyContent: 'flex-end', gap: 10, flexDirection: isMobile ? 'column' : 'row' }}>
            <button
              onClick={() => this.setState({ addFolderDialog: null })}
              disabled={busy}
              style={{
                border: 'none', background: clay.surface, borderRadius: clay.radiusButton,
                minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 18px', fontSize: 13, fontWeight: 600, color: clay.textMuted,
                cursor: busy ? 'not-allowed' : 'pointer', boxShadow: clay.shadowRaised,
              }}
            >
              Cancel
            </button>
            <button
              onClick={() => void this._submitAddFolder()}
              disabled={busy || !name.trim()}
              style={{
                border: 'none', background: clay.accentGradient, color: '#fff', borderRadius: clay.radiusButton,
                minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', fontSize: 13, fontWeight: 700,
                cursor: (busy || !name.trim()) ? 'not-allowed' : 'pointer', boxShadow: clay.shadowButton,
                opacity: (busy || !name.trim()) ? 0.6 : 1,
              }}
            >
              {busy ? 'Creating…' : addFolderDialog.asVessel ? 'Continue to Add Vessel' : '+ Create Folder'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  public _openFolderDeleteDialog = (options: {
    folderId: string;
    folderName: string;
    folderPath: string;
    vesselName: string;
    mainFolder?: string;
    subfolderNames?: string[];
  }): void => {
    const { folderId, folderName, folderPath, vesselName, mainFolder, subfolderNames = [] } = options;
    const activeMain = mainFolder || this.state.docMainFolder || 'Technical & Crewing';

    const availableFolders: Array<{ id: string; name: string; path: string; isCurrent?: boolean }> = [];

    const currentPathParts = (folderPath || '').split('>').map(s => s.trim()).filter(Boolean);
    const currentPathSlash = currentPathParts.join('/');

    const findFolderIdByPathOrName = (name: string): string | undefined => {
      const targetPathEndsWith = `${currentPathSlash}/${name}`.toLowerCase();
      const targetName = name.toLowerCase();

      // 1. Prefer exact path match in live folder map.
      for (const [id, node] of Array.from(this.state.spoFolderMap.entries())) {
        const norm = ((node?.serverRelativePath || '')).trim().replace(/^\/+|\/+$/g, '').toLowerCase();
        if (!norm) continue;
        if (norm === targetPathEndsWith || norm.endsWith(`/${targetPathEndsWith}`)) {
          return id;
        }
      }

      // 2. Exact child name under parent path match in live folder map.
      for (const [id, node] of Array.from(this.state.spoFolderMap.entries())) {
        const norm = ((node?.serverRelativePath || '')).trim().replace(/^\/+|\/+$/g, '').toLowerCase();
        if (!norm) continue;
        if (node.name?.toLowerCase() === targetName && currentPathSlash && norm.includes(currentPathSlash.toLowerCase())) {
          return id;
        }
      }

      // 3. Search rows — normalize subFolderPath separators (rows use ' > ', currentPathSlash uses '/').
      const currentPathNorm = currentPathSlash.toLowerCase();
      const rowMatch = this.state.rows.find(r => {
        const rowVessel = (r.vesselName || '').toLowerCase();
        const rowSub = (r.subCategory || '').toLowerCase();
        const rowCat = (r.category || '').toLowerCase();
        // Normalize row path: replace ' > ' with '/' for comparison
        const rowPath = (r.subFolderPath || '').replace(/\s*>\s*/g, '/').toLowerCase();
        const vesselMatches = !vesselName || rowVessel === vesselName.toLowerCase();
        const nameMatches = rowSub === targetName || rowCat === targetName;
        // Accept if path includes current context OR if just vessel+name matches (no path context)
        const pathMatches = !currentPathNorm || rowPath.includes(currentPathNorm);
        const hasRealId = !!r.uploadFolderId && !r.uploadFolderId.includes('/');
        return vesselMatches && nameMatches && pathMatches && hasRealId;
      });
      if (rowMatch?.uploadFolderId) return rowMatch.uploadFolderId;

      // 4. Loose fallback: match by name + vessel only (no path constraint) — picks up
      //    rows where subFolderPath contains the uploaded folder name itself.
      const looseMatch = this.state.rows.find(r => {
        const rowVessel = (r.vesselName || '').toLowerCase();
        const rowSub = (r.subCategory || '').toLowerCase();
        const rowCat = (r.category || '').toLowerCase();
        const vesselMatches = !vesselName || rowVessel === vesselName.toLowerCase();
        const nameMatches = rowSub === targetName || rowCat === targetName;
        const hasRealId = !!r.uploadFolderId && !r.uploadFolderId.includes('/');
        return vesselMatches && nameMatches && hasRealId;
      });
      return looseMatch?.uploadFolderId;
    };

    // Add current folder
    if (folderName && !/^(vessels_root|specific_vessels|kaizen_root)$/i.test(folderName)) {
      availableFolders.push({
        id: folderId || folderName,
        name: folderName,
        path: folderPath || folderName,
        isCurrent: true,
      });
    }

    // Add child subfolders if any
    for (const sf of subfolderNames) {
      if (sf && sf !== folderName && !availableFolders.some(f => f.name.toLowerCase() === sf.toLowerCase())) {
        const resolvedId = findFolderIdByPathOrName(sf) || sf;
        availableFolders.push({
          id: resolvedId,
          name: sf,
          path: folderPath ? `${folderPath} > ${sf}` : sf,
          isCurrent: false,
        });
      }
    }

    if (availableFolders.length === 0) return;

    this.setState({
      folderDeleteDialog: {
        currentFolderId: folderId,
        currentFolderName: folderName,
        currentFolderPath: folderPath,
        vesselName: vesselName,
        mainFolder: activeMain,
        availableFolders,
        selectedFolderId: availableFolders[0].id,
        busy: false,
        error: null,
      },
    });
  };

  public _deleteSelectedFolder = async (): Promise<void> => {
    const dialog = this.state.folderDeleteDialog;
    if (!dialog || !dialog.selectedFolderId) return;

    this.setState(prev => ({
      folderDeleteDialog: prev.folderDeleteDialog ? { ...prev.folderDeleteDialog, busy: true, error: null } : null,
    }));

    const { selectedFolderId, availableFolders, vesselName, mainFolder } = dialog;
    const target = availableFolders.find(f => f.id === selectedFolderId) || availableFolders[0];
    if (!target) return;

    const { graphClient, siteId, driveId } = this.props;
    const activeMain = mainFolder || this.state.docMainFolder || 'Technical & Crewing';
    const MAIN_FOLDERS_LIST = [
      'Technical & Crewing', 'Commercial & Chartering', 'Commercial', 'Insurance',
      'Kaizen - Knowledge Bank', 'Knowledge Bank', 'Technical and Crewing', 'Commercial and Chartering'
    ];

    // Construct canonical target SharePoint path
    const rawSegments = (target.path || target.name || '')
      .split(/[>/]/)
      .map(s => s.trim())
      .filter(Boolean);

    const isKaizen = vesselName === 'Kaizen - Knowledge Bank' ||
      rawSegments.some(s => s.toLowerCase() === 'kaizen - knowledge bank');

    const isCommon = vesselName === 'Common for all vessels' ||
      vesselName === 'Common for all ships' ||
      rawSegments.some(s => /^common for all (vessels|ships)$/i.test(s));

    let canonicalSharePointPath: string;
    if (isKaizen) {
      const kaizenParts = rawSegments.filter(s => s.toLowerCase() !== 'kaizen - knowledge bank');
      canonicalSharePointPath = ['Kaizen - Knowledge Bank', ...kaizenParts].join('/');
    } else if (isCommon) {
      const commonParts = rawSegments.filter(s =>
        !/^common for all (vessels|ships)$/i.test(s) &&
        !MAIN_FOLDERS_LIST.some(mf => mf.toLowerCase() === s.toLowerCase())
      );
      canonicalSharePointPath = [activeMain, 'Common for all ships', ...commonParts].join('/');
    } else {
      const effectiveVessel = vesselName && vesselName !== 'all' ? vesselName : (rawSegments[0] || '');
      const subTree = rawSegments.filter(s =>
        s.toLowerCase() !== effectiveVessel.toLowerCase() &&
        !MAIN_FOLDERS_LIST.some(mf => mf.toLowerCase() === s.toLowerCase()) &&
        !/^(vessels|specific vessels|documents|root)$/i.test(s)
      );
      canonicalSharePointPath = [activeMain, effectiveVessel, ...subTree].filter(Boolean).join('/');
    }

    const encodePath = (p: string) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

    let deletedSPO = false;
    let spoItemId = target.id;
    let recycleBinItemId: string | undefined;

    // 1. Delete in SharePoint Online via Graph API (moves to SPO Recycle Bin)
    if (graphClient && siteId && driveId) {
      try {
        let driveItemId: string | null = null;
        const hasLikelyDriveItemId = !!target.id &&
          target.id !== target.name &&
          !target.id.includes('/') &&
          !target.id.includes('>') &&
          !target.id.includes('\\') &&
          !target.id.startsWith('sf_') &&
          !/^\d+$/.test(target.id);

        if (hasLikelyDriveItemId) {
          driveItemId = target.id;
        } else {
          // Check spoFolderMap first
          const targetNormPath = canonicalSharePointPath.toLowerCase();
          for (const [id, node] of Array.from(this.state.spoFolderMap.entries())) {
            const nodeNorm = (node?.serverRelativePath || '').trim().replace(/^\/+|\/+$/g, '').toLowerCase();
            if (nodeNorm && (nodeNorm === targetNormPath || nodeNorm.endsWith(`/${targetNormPath}`))) {
              driveItemId = id;
              break;
            }
          }

          // If not in map, look up drive item ID by path in Graph
          if (!driveItemId) {
            const candidatePaths = [
              canonicalSharePointPath,
              (target.path || '').replace(/\s*>\s*/g, '/'),
              canonicalSharePointPath.replace(new RegExp(`^${activeMain}/`, 'i'), ''),
            ].filter(Boolean);

            for (const cPath of candidatePaths) {
              try {
                const item = await graphClient
                  .api(`/sites/${siteId}/drives/${driveId}/root:/${encodePath(cPath)}?$select=id,name,webUrl`)
                  .get();
                if (item?.id) {
                  driveItemId = item.id;
                  break;
                }
              } catch {
                // Continue to next candidate path
              }
            }
          }
        }

        // Attempt direct-ID delete; if 404, fall back to path-based resolution
        if (driveItemId) {
          let directDeleteOk = false;
          try {
            await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${driveItemId}`).delete();
            directDeleteOk = true;
          } catch (directErr: any) {
            const is404 = (directErr?.statusCode === 404) || (String(directErr?.message || '').includes('itemNotFound'));
            if (!is404) throw directErr;
            console.warn(`[VesselDMS] Direct ID delete 404 for "${target.name}" (${driveItemId}), trying path lookup…`);
          }

          if (!directDeleteOk) {
            // Fall back to path-based resolution
            driveItemId = null;
            const fallbackPaths = [
              canonicalSharePointPath,
              (target.path || '').replace(/\s*>\s*/g, '/'),
              canonicalSharePointPath.replace(new RegExp(`^${activeMain}/`, 'i'), ''),
            ].filter(Boolean);

            for (const cPath of fallbackPaths) {
              try {
                const item = await graphClient
                  .api(`/sites/${siteId}/drives/${driveId}/root:/${encodePath(cPath)}?$select=id,name,webUrl`)
                  .get();
                if (item?.id) {
                  driveItemId = item.id;
                  break;
                }
              } catch { /* try next */ }
            }

            if (driveItemId) {
              await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${driveItemId}`).delete();
            } else {
              console.warn(`[VesselDMS] _deleteSelectedFolder: could not resolve "${target.name}" via any path`);
            }
          }

          if (driveItemId) {
            spoItemId = driveItemId;
            this._appDeletedItemIds.add(driveItemId);
            deletedSPO = true;
            console.log(`[VesselDMS] _deleteSelectedFolder: deleted SPO folder "${target.name}" (${driveItemId})`);

            // Capture recycle bin item ID AFTER delete
            if (this.props.siteUrl) {
              try {
                await new Promise(res => setTimeout(res, 1200));
                const rbRes = await fetch(
                  `${this.props.siteUrl}/_api/site/RecycleBin?$filter=LeafName eq '${encodeURIComponent(target.name)}'&$select=Id,LeafName&$top=10`,
                  { headers: { Accept: 'application/json;odata=nometadata' } }
                );
                if (rbRes.ok) {
                  const rbData = await rbRes.json();
                  const match = (rbData?.value ?? []).find((r: any) => (r.LeafName || '').toLowerCase() === target.name.toLowerCase());
                  if (match?.Id) recycleBinItemId = match.Id;
                }
              } catch { /* non-critical */ }
            }
          }
        }
      } catch (graphErr) {
        console.warn('[VesselDMS] _deleteSelectedFolder Graph delete error:', graphErr);
      }
    }


    // 2. Notify backend (pass driveItemId if available, or path/target.id)
    const backendFolderKey = spoItemId || canonicalSharePointPath || target.id;
    const folderReasonParam = dialog.reason?.trim() ? `&reason=${encodeURIComponent(dialog.reason.trim())}` : '';
    try {
      if (backendFolderKey && !backendFolderKey.startsWith('sf_')) {
        await fetch(`${this._base()}/api/folders/${encodeURIComponent(backendFolderKey)}?folder_name=${encodeURIComponent(target.name)}${folderReasonParam}`, {
          method: 'DELETE', headers: this._headers(),
        }).catch(() => undefined);
      }
    } catch { /* non-critical */ }

    // 3. Create Recycle Bin entry
    const activeSiteObj = (this.state.documentSites || []).find(s => s.site_key === this.state.activeDocumentSite);
    const folderVessel = this.state.vessels.find(v => (v.name || '').toLowerCase() === (vesselName || '').toLowerCase());
    const siteNames = (folderVessel?.provisioned_site_ids || []).map(sk => {
      const matched = (this.state.documentSites || []).find(s => s.site_key === sk);
      return matched?.sp_site_name || sk;
    }).filter(Boolean);
    const resolvedSiteName = activeSiteObj?.sp_site_name || (siteNames.length > 0 ? siteNames.join(', ') : (this.props.siteUrl || 'SharePoint'));

    const deletedNode: import('./types/ui').DeletedNode = {
      id: spoItemId || target.id || `folder_${Date.now()}`,
      name: target.name,
      kind: 'folder',
      item_type: 'folder',
      main_folder: activeMain,
      original_path: canonicalSharePointPath,
      vessel_name: vesselName,
      category: target.name,
      sub_category: target.name,
      deleted_at: new Date().toISOString(),
      in_spo_recycle_bin: deletedSPO,
      recycle_bin_item_id: recycleBinItemId,
      site_name: resolvedSiteName,
      site_key: activeSiteObj?.site_key || this.state.activeDocumentSite || undefined,
    };
    // Graph-direct delete above already moved this to SharePoint's own
    // Recycle Bin when deletedSPO is true — the backend call in step 2 is
    // a best-effort notify, so log deletion attribution here regardless of
    // which path actually performed the delete.
    this._logDeletion({
      item_type: 'folder',
      drive_item_id: spoItemId || target.id || null,
      name: target.name,
      original_path: canonicalSharePointPath,
      site_name: resolvedSiteName,
      site_key: activeSiteObj?.site_key || this.state.activeDocumentSite || null,
      reason: dialog.reason || null,
    });

    // 4. Update state: add to recycleBin, remove from spoFolderMap, rows, and uploadedFilesByFolder
    const targetNorm = target.name.toLowerCase();
    if (target.id) this._appDeletedItemIds.add(target.id);
    if (spoItemId) this._appDeletedItemIds.add(spoItemId);

    this.setState(prev => {
      const nextMap = new Map(prev.spoFolderMap);
      if (spoItemId) nextMap.delete(spoItemId);
      if (target.id) nextMap.delete(target.id);

      // Clean up children in all parent nodes in spoFolderMap
      for (const [id, node] of Array.from(nextMap.entries())) {
        if (node?.children && Array.isArray(node.children)) {
          const filtered = node.children.filter(c =>
            c.id !== spoItemId && c.id !== target.id && (c.name || '').toLowerCase() !== targetNorm
          );
          if (filtered.length !== node.children.length) {
            nextMap.set(id, { ...node, children: filtered });
          }
        }
      }

      // Clean up uploadedFilesByFolder
      const nextUploadedByFolder = { ...prev.uploadedFilesByFolder };
      Object.keys(nextUploadedByFolder).forEach(k => {
        const kLow = k.toLowerCase();
        if (
          kLow === targetNorm ||
          kLow.endsWith(`/${targetNorm}`) ||
          kLow.endsWith(`>${targetNorm}`) ||
          kLow.endsWith(`> ${targetNorm}`) ||
          kLow.includes(`>${targetNorm}>`) ||
          kLow.includes(`> ${targetNorm} >`) ||
          (target.id && k === target.id) ||
          (spoItemId && k === spoItemId) ||
          (canonicalSharePointPath && kLow === canonicalSharePointPath.toLowerCase())
        ) {
          delete nextUploadedByFolder[k];
        }
      });

      const targetPathNorm = (target.path || '').replace(/\s*>\s*/g, ' > ').trim().toLowerCase();
      const targetCanonicalLow = canonicalSharePointPath.toLowerCase();
      const nextRows = prev.rows.filter(r => {
        const rowPath = (r.subFolderPath || '').replace(/\s*>\s*/g, ' > ').trim().toLowerCase();
        const rowPathSlash = (r.subFolderPath || '').replace(/\s*>\s*/g, '/').trim().toLowerCase();
        if (targetPathNorm && (rowPath === targetPathNorm || rowPath.startsWith(`${targetPathNorm} > `))) return false;
        if (targetCanonicalLow && (rowPathSlash === targetCanonicalLow || rowPathSlash.startsWith(`${targetCanonicalLow}/`))) return false;
        if ((r.subCategory || '').toLowerCase() === targetNorm) return false;
        if ((r.category || '').toLowerCase() === targetNorm && (target as any).isCategoryLevel) return false;
        if (r.uploadFolderId && (r.uploadFolderId === target.id || r.uploadFolderId === spoItemId || this._appDeletedItemIds.has(r.uploadFolderId))) return false;
        return true;
      });

      this._persistUploadCache(nextRows, nextUploadedByFolder);

      return {
        recycleBin: [deletedNode, ...prev.recycleBin.filter(r => r.id !== deletedNode.id)],
        spoFolderMap: nextMap,
        uploadedFilesByFolder: nextUploadedByFolder,
        rows: nextRows,
        folderDeleteDialog: null,
      };
    });

    // 5. If the user was inside the deleted folder, navigate back up
    if (target.isCurrent) {
      const nextStack = this.state.folderPathStack.slice(0, -1);
      this._pushFolderNav(nextStack, this.state.docMainFolder);
    }

    void this._syncScheduler?.triggerNow().catch(() => undefined);
  };

  public _renderFolderDeleteDialog(): React.ReactElement | null {
    const { folderDeleteDialog } = this.state;
    if (!folderDeleteDialog) return null;
    const { availableFolders, selectedFolderId, busy, error } = folderDeleteDialog;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

    const selectedFolder = availableFolders.find(f => f.id === selectedFolderId) || availableFolders[0];

    return (
      <div style={{
        position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.6)', zIndex: 100001,
        backdropFilter: 'blur(4px)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 20,
      }}>
        <div style={{
          background: 'var(--vdms-surface)', borderRadius: 16, width: isMobile ? '95vw' : 480, maxWidth: '95vw', maxHeight: '90vh', overflow: 'hidden',
          boxShadow: '0 20px 60px rgba(15,23,42,0.3)', border: '1px solid #fee2e2',
        }}>
          {/* Header */}
          <div style={{ padding: '20px 24px 16px', background: '#fff1f2', borderBottom: '1px solid #fecdd3' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <div style={{ width: 42, height: 42, borderRadius: 10, background: '#ffe4e6', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>
                <Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 22 }} />
              </div>
              <div>
                <div style={{ fontSize: 17, fontWeight: 700, color: '#9f1239' }}>
                  Move Folder to Recycle Bin
                </div>
                <div style={{ fontSize: 12, color: '#be123c', marginTop: 2 }}>
                  Select the uploaded folder to remove from SharePoint Online
                </div>
              </div>
            </div>
          </div>

          {/* Body */}
          <div style={{ padding: '20px 24px' }}>
            {error && (
              <div style={{ background: '#fee2e2', color: '#991b1b', padding: '10px 14px', borderRadius: 8, fontSize: 12, marginBottom: 14 }}>
                <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {error}
              </div>
            )}

            <div style={{ fontSize: 13, color: 'var(--vdms-text)', fontWeight: 600, marginBottom: 8 }}>
              Select folder to delete:
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: 200, overflowY: 'auto', marginBottom: 16 }}>
              {availableFolders.map(folder => {
                const isChecked = folder.id === selectedFolderId;
                return (
                  <label
                    key={folder.id}
                    onClick={() => this.setState(prev => ({
                      folderDeleteDialog: prev.folderDeleteDialog ? { ...prev.folderDeleteDialog, selectedFolderId: folder.id } : null,
                    }))}
                    style={{
                      display: 'flex', alignItems: 'center', gap: 12, padding: '12px 14px',
                      borderRadius: 10, border: `1.5px solid ${isChecked ? '#e11d48' : 'var(--vdms-border)'}`,
                      background: isChecked ? '#fff1f2' : 'var(--vdms-surface-alt)', cursor: 'pointer',
                      transition: 'all 0.15s',
                    }}
                  >
                    <input
                      type="radio"
                      name="selectedFolder"
                      checked={isChecked}
                      onChange={() => {}}
                      style={{ width: 18, height: 18, accentColor: '#e11d48', cursor: 'pointer' }}
                    />
                    <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 20 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: 'var(--vdms-text)' }}>
                        {folder.name} {folder.isCurrent ? <span style={{ fontSize: 11, color: '#e11d48', fontWeight: 600 }}>(Current Folder)</span> : ''}
                      </div>
                      <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={folder.path}>
                        {folder.path}
                      </div>
                    </div>
                  </label>
                );
              })}
            </div>

            <div style={{ background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-border)', borderRadius: 10, padding: 12, fontSize: 12, color: 'var(--vdms-text-muted)', lineHeight: 1.5, marginBottom: 16 }}>
              ℹ️ Moving <strong style={{ color: 'var(--vdms-text)' }}>{selectedFolder?.name}</strong> to the Recycle Bin will also delete its contained files in SharePoint Online. You can restore this folder anytime from the <strong style={{ color: 'var(--vdms-text)' }}>Recycle Bin</strong> page.
            </div>

            <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: 'var(--vdms-text-muted)', textTransform: 'uppercase', marginBottom: 6 }}>
              Reason for deletion (optional)
            </label>
            <textarea
              value={folderDeleteDialog.reason || ''}
              onChange={e => this.setState(prev => (
                prev.folderDeleteDialog ? { folderDeleteDialog: { ...prev.folderDeleteDialog, reason: e.target.value } } : null as any
              ))}
              disabled={busy}
              placeholder="Optional — why is this being deleted?"
              rows={2}
              style={{
                width: '100%', boxSizing: 'border-box', borderRadius: 8, border: '1px solid var(--vdms-border)',
                padding: '8px 10px', fontSize: 13, fontFamily: 'inherit', resize: 'vertical',
                background: 'var(--vdms-surface)', color: 'var(--vdms-text)',
              }}
            />
          </div>

          {/* Footer */}
          <div style={{ padding: '14px 24px 20px', background: 'var(--vdms-surface-alt)', borderTop: '1px solid var(--vdms-border)', display: 'flex', justifyContent: 'flex-end', gap: 10, flexDirection: isMobile ? 'column' : 'row' }}>
            <button
              onClick={() => this.setState({ folderDeleteDialog: null })}
              disabled={busy}
              style={{
                border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface)', borderRadius: 8,
                minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 18px', fontSize: 13, fontWeight: 600, color: 'var(--vdms-text-secondary)',
                cursor: busy ? 'not-allowed' : 'pointer',
              }}
            >
              Cancel
            </button>
            <button
              onClick={() => void this._deleteSelectedFolder()}
              disabled={busy}
              style={{
                border: 'none', background: '#e11d48', color: '#fff', borderRadius: 8,
                minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', fontSize: 13, fontWeight: 700,
                cursor: busy ? 'not-allowed' : 'pointer', boxShadow: '0 2px 6px rgba(225,29,72,0.3)',
                opacity: busy ? 0.65 : 1, display: 'inline-flex', alignItems: 'center', gap: 6,
              }}
            >
              {busy ? 'Moving to Recycle Bin…' : <><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 13 }} /> Move to Recycle Bin</>}
            </button>
          </div>
        </div>
      </div>
    );
  }

  public _uploadSuccessTimer: ReturnType<typeof setInterval> | null = null;

  public _startUploadSuccessTimer = (): void => {
    if (this._uploadSuccessTimer) clearInterval(this._uploadSuccessTimer);
    this._uploadSuccessTimer = setInterval(() => {
      this.setState(prev => {
        if (!prev.uploadSuccessPopup) { clearInterval(this._uploadSuccessTimer!); return null as any; }
        const next = prev.uploadSuccessPopup.secondsLeft - 1;
        if (next <= 0) { clearInterval(this._uploadSuccessTimer!); return { uploadSuccessPopup: null }; }
        return { uploadSuccessPopup: { ...prev.uploadSuccessPopup, secondsLeft: next } };
      });
    }, 1000);
  };

  private _formatUploadSize(file: File): string {
    if (!file || typeof file.size !== 'number') return '—';
    if (file.size < 1024) return `${file.size} B`;
    if (file.size < 1024 * 1024) return `${(file.size / 1024).toFixed(1)} KB`;
    return `${(file.size / (1024 * 1024)).toFixed(2)} MB`;
  }

  private _uploadVisibilityKeys(row: GroupedRow): string[] {
    const keys = new Set<string>();
    const add = (value?: string | null): void => {
      const v = String(value || '').trim();
      if (!v) return;
      keys.add(v);
      keys.add(v.toLowerCase());
    };
    add(row.groupKey);
    add(row.uploadFolderId);
    add(row.subFolderPath);
    add(row.subCategory);
    add(row.category);
    if (row.vesselName && row.subCategory) add(`${row.vesselName} > ${row.subCategory}`);
    if (row.vesselName && row.group && row.subCategory) add(`${row.vesselName} > ${row.group} > ${row.subCategory}`);
    return Array.from(keys);
  }

  public _addUploadingFilePlaceholder(keys: string[], file: File): string {
    const tempId = `file_uploading_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
    const placeholder = {
      id: tempId,
      name: file.name,
      size: this._formatUploadSize(file),
      date: 'Uploading...',
      uploading: true,
      pending: false,
      uploadedAt: Date.now(),
    };
    this.setState(prev => {
      const next = { ...prev.uploadedFilesByFolder };
      (keys || []).forEach(key => {
        const k = (key || '').trim();
        if (!k) return;
        const current = Array.isArray(next[k]) ? next[k] : [];
        const pruned = current.filter(f => !((f as any).uploading && f.name === file.name));
        next[k] = [...pruned, placeholder];
      });
      return { uploadedFilesByFolder: next };
    });
    return tempId;
  }

  public _clearUploadingFilePlaceholder(keys: string[], tempId: string): void {
    this.setState(prev => {
      const next = { ...prev.uploadedFilesByFolder };
      let changed = false;
      (keys || []).forEach(key => {
        const k = (key || '').trim();
        if (!k || !Array.isArray(next[k])) return;
        const before = next[k];
        const after = before.filter(f => (f as any).id !== tempId);
        if (after.length !== before.length) {
          next[k] = after;
          changed = true;
        }
      });
      return changed ? { uploadedFilesByFolder: next } : null;
    });
  }

  public _replaceUploadingFilePlaceholder(
    keys: string[],
    tempId: string,
    file: File,
    finalInfo: { id?: string | null; pending?: boolean; date?: string; uploadedAt?: number }
  ): void {
    const finalEntry = {
      id: finalInfo.id || tempId,
      name: file.name,
      size: this._formatUploadSize(file),
      date: finalInfo.date || 'Just now',
      pending: Boolean(finalInfo.pending),
      uploading: false,
      uploadedAt: finalInfo.uploadedAt || Date.now(),
    };

    this.setState(prev => {
      const next = { ...prev.uploadedFilesByFolder };
      (keys || []).forEach(key => {
        const k = (key || '').trim();
        if (!k) return;
        const current = Array.isArray(next[k]) ? next[k] : [];
        const pruned = current.filter(f => (f as any).id !== tempId && f.name !== file.name);
        next[k] = [...pruned, finalEntry];
      });
      return { uploadedFilesByFolder: next };
    });
  }

  public _handleUpload = async (row: GroupedRow, files: File[]): Promise<void> => {
    if (!files.length) return;
    this.setState({ uploadingGroupKey: row.groupKey, uploadError: null, uploadInfo: `Uploading ${files.length === 1 ? files[0].name : `${files.length} files`}…` });
    let done = 0, failed = 0, pending = 0;
    let lastFolderId: string | null = null;
    let lastIsGraphUpload = false;
    let lastWebUrl = '';
    let lastDestPath = row.subFolderPath || row.subCategory || row.category;
    const successfulUploadEntries: VesselSuggestionUploadEntry[] = [];
    const uploadKeys = this._uploadVisibilityKeys(row);
    for (const file of files) {
      const tempUploadId = this._addUploadingFilePlaceholder(uploadKeys, file);
      try {
        const result = await this._uploadFileToFolder(
          row.uploadFolderId,
          row.subFolderPath,
          row.vesselName,
          file,
          row.monthDriven
        );
        if (result.statusPending) pending++;
        else done++;
        this._replaceUploadingFilePlaceholder(uploadKeys, tempUploadId, file, {
          id: result.fileId,
          pending: result.statusPending,
          date: result.statusPending ? 'Queued for approval' : 'Just now',
          uploadedAt: Date.now(),
        });
        lastFolderId = result.folderId || lastFolderId;
        lastIsGraphUpload = result.isGraphUpload || lastIsGraphUpload;
        lastWebUrl = result.webUrl || lastWebUrl;
        successfulUploadEntries.push({
          filename: file.name,
          drive_item_id: result.fileId || null,
          source_subfolder_path: row.subFolderPath || null,
          source_vessel_name: row.vesselName || null,
          uploaded_at: Date.now(),
        });
      } catch (error) {
        failed++;
        this._clearUploadingFilePlaceholder(uploadKeys, tempUploadId);
        const detail = error instanceof Error ? error.message : 'Upload failed';
        this.setState({ uploadError: detail });
      }
    }
    this.setState({ uploadingGroupKey: null, uploadInfo: null, uploadError: failed > 0 ? `❌ ${failed} file(s) failed.` : null });
    if (done > 0 || pending > 0) {
      const firstName = files[0]?.name || 'file';
      this.setState({
        uploadSuccessPopup: {
          fileName: firstName,
          destinationPath: lastDestPath,
          webUrl: lastWebUrl,
          isPending: pending > 0 && done === 0,
          secondsLeft: 10,
          unidentifiedFiles: [],
        },
      });
      this._startUploadSuccessTimer();
    }
    if (done > 0 || pending > 0) {
      const refreshFolderId = lastFolderId || row.uploadFolderId;
      // Use isGraphUpload flag (true = folderId is a real Graph drive item ID),
      // not the presence of folderId (which could be a backend DB ID).
      const isRealGraphId = lastIsGraphUpload;
      // Single refresh - _refreshFolderFiles now handles Graph consistency window
      // by preserving local uploads when Graph returns empty.
      void this._refreshFolderFiles(refreshFolderId, row.groupKey, true, isRealGraphId).catch(() => undefined);
      // Reconcile with SharePoint after the optimistic file list is visible.
      // Running delta immediately makes the whole documents view enter its
      // loading state and can temporarily replace the current folder contents.
      window.setTimeout(() => {
        void this._syncScheduler?.triggerNow().catch(() => undefined);
      }, 2000);
    }
  };

  public _renderUploadSuccessPopup(): React.ReactElement | null {
    const p = this.state.uploadSuccessPopup;
    if (!p) return null;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
    return (
      <div style={{
        position: 'fixed', bottom: isMobile ? 10 : 28, right: isMobile ? 10 : 28, zIndex: 100002,
        background: 'var(--vdms-surface)', borderRadius: 16, boxShadow: '0 8px 40px rgba(0,0,0,0.18)',
        border: '1.5px solid #86efac', padding: isMobile ? '12px' : '20px 24px 18px', width: isMobile ? 'calc(100vw - 20px)' : 'auto', minWidth: isMobile ? 0 : 340, maxWidth: isMobile ? 'calc(100vw - 20px)' : 420,
        fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", animation: 'slideInRight 0.3s ease',
      }}>
        <style>{`
          @keyframes slideInRight { from { opacity:0; transform:translateX(40px); } to { opacity:1; transform:translateX(0); } }
          @keyframes countdownShrink { from { width:100%; } to { width:0%; } }
        `}</style>
        {/* Header row */}
        <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12, marginBottom: 10 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{
              width: 36, height: 36, borderRadius: 10,
              background: p.isPending ? '#fef3c7' : '#dcfce7',
              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0,
            }}>
              {p.isPending ? <Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 18 }} /> : <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 18 }} />}
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)' }}>
                {p.isPending ? 'Submitted for Approval' : 'Upload Successful!'}
              </div>
              <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                Auto-closes in {p.secondsLeft}s
              </div>
            </div>
          </div>
          <button
            onClick={() => { clearInterval(this._uploadSuccessTimer!); this.setState({ uploadSuccessPopup: null }); }}
            style={{ border: 'none', background: 'none', cursor: 'pointer', color: 'var(--vdms-text-faint)', minHeight: 44, minWidth: 44, fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}
            title="Close"
          ><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 18 }} /></button>
        </div>
        {/* File info */}
        <div style={{ background: 'var(--vdms-surface-alt)', borderRadius: 8, padding: '8px 12px', marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginBottom: 3 }}><Icon iconName="Page" aria-hidden="true" style={{ fontSize: 12 }} /> File</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: 'var(--vdms-text)', wordBreak: 'break-all' }}>{p.fileName}</div>
          <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)', marginTop: 6 }}><Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 11 }} /> Path</div>
          <div style={{ fontSize: 12, color: 'var(--vdms-text)', marginTop: 2, wordBreak: 'break-all' }}>{p.destinationPath}</div>
          {Array.isArray(p.unidentifiedFiles) && p.unidentifiedFiles.length > 0 && (
            <div style={{ marginTop: 10, background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 8, padding: '8px 10px' }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: '#9a3412' }}><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> Vessel name not identified files</div>
              <div style={{ marginTop: 4, fontSize: 11, color: '#9a3412' }}>
                Sent to Templates & OCR Need Review for manual classification.
              </div>
              <div style={{ marginTop: 6, display: 'flex', flexDirection: 'column', gap: 3 }}>
                {p.unidentifiedFiles.slice(0, 4).map((name, idx) => (
                  <div key={`${name}_${idx}`} style={{ fontSize: 11, color: '#7c2d12', wordBreak: 'break-all' }}>• {name}</div>
                ))}
                {p.unidentifiedFiles.length > 4 && (
                  <div style={{ fontSize: 11, color: '#9a3412' }}>+{p.unidentifiedFiles.length - 4} more file(s)</div>
                )}
              </div>
            </div>
          )}
        </div>
        {/* Action buttons */}
        <div style={{ display: 'flex', gap: 8, flexDirection: isMobile ? 'column' : 'row' }}>
          <button
            onClick={() => {
              clearInterval(this._uploadSuccessTimer!);
              this.setState({ uploadSuccessPopup: null });
              this._toggleAlertBell();
            }}
            style={{
              flex: 1, background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: 8,
              minHeight: 44, padding: '8px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer',
              boxShadow: clay.shadowButton,
            }}
          >
            View Alerts
          </button>
          {p.webUrl && (
            <a
              href={p.webUrl}
              target="_blank"
              rel="noopener noreferrer"
              style={{
                flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                background: clay.accentGradient, color: '#fff', borderRadius: 8,
                minHeight: 44, padding: '8px 12px', fontSize: 12, fontWeight: 600, textDecoration: 'none',
                cursor: 'pointer', boxShadow: clay.shadowButton,
              }}
            >
              <Icon iconName="Link" aria-hidden="true" style={{ fontSize: 12 }} /> Open in SharePoint
            </a>
          )}
          <button
            onClick={() => { clearInterval(this._uploadSuccessTimer!); this.setState({ uploadSuccessPopup: null }); }}
            style={{
              flex: 1, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)', border: 'none', borderRadius: 8,
              minHeight: 44, padding: '8px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            Stay here
          </button>
        </div>
        {/* Countdown progress bar */}
        <div style={{ marginTop: 12, height: 3, background: 'var(--vdms-surface-alt)', borderRadius: 2, overflow: 'hidden' }}>
          <div style={{
            height: '100%', borderRadius: 2,
            background: p.isPending ? '#f59e0b' : '#16a34a',
            width: `${(p.secondsLeft / 10) * 100}%`,
            transition: 'width 1s linear',
          }} />
        </div>
      </div>
    );
  }

  public _renderUploadNavigationPrompt(): React.ReactElement | null {
    // Disabled per user request: we do not want review & ocr
    return null;
  }

  public _getGrouped(): GroupedRow[] {
    const { rows, textFilter, vesselFilter, groupFilter, catFilter, sort } = this.state;
    const q = textFilter.toLowerCase();
    let f = rows;
    if (vesselFilter !== 'all') f = f.filter(r => r.vesselName === vesselFilter);
    if (groupFilter !== 'all') f = f.filter(r => r.group === groupFilter);
    if (catFilter !== 'all') f = f.filter(r => r.category === catFilter);
    if (q) f = f.filter(r => [r.vesselName, r.group, r.category, r.subFolderPath, r.fileName ?? ''].some(s => s.toLowerCase().indexOf(q) !== -1));

    const map = new Map<string, GroupedRow>();
    for (const row of f) {
      const ex = map.get(row.groupKey);
      if (!ex) map.set(row.groupKey, { srNo: row.srNo, vesselName: row.vesselName, group: row.group, category: row.category, subCategory: row.subCategory || row.category, subFolderPath: row.subFolderPath, groupKey: row.groupKey, uploadFolderId: row.uploadFolderId, monthDriven: row.monthDriven, canUpload: row.canUpload, files: row.fileId && row.fileName ? [{ id: row.fileId, name: row.fileName }] : [] });
      else if (row.fileId && row.fileName) ex.files.push({ id: row.fileId, name: row.fileName });
    }

    const grouped = Array.from(map.values());
    if (sort === 'name_asc') grouped.sort((a, b) => a.vesselName.localeCompare(b.vesselName));
    else if (sort === 'name_desc') grouped.sort((a, b) => b.vesselName.localeCompare(a.vesselName));
    return grouped;
  }

  // ── Vessel Suggestions Module ──────────────────────────────────────────────

  /** Known shipyard keyword → canonical display name map. */
  private static readonly SHIPYARD_MAP: Array<{ keywords: RegExp; name: string }> = [
    { keywords: /hyundai/i, name: 'Hyundai Heavy Industries' },
    { keywords: /samsung/i, name: 'Samsung Heavy Industries' },
    { keywords: /daewoo|dsme/i, name: 'Daewoo Shipbuilding & Marine Engineering' },
    { keywords: /imabari/i, name: 'Imabari Shipbuilding' },
    { keywords: /oshima/i, name: 'Oshima Shipbuilding' },
    { keywords: /tsuneishi/i, name: 'Tsuneishi Shipbuilding' },
    { keywords: /hudong/i, name: 'Hudong-Zhonghua Shipbuilding' },
    { keywords: /yangzijiang/i, name: 'Yangzijiang Shipbuilding' },
    { keywords: /cosco/i, name: 'COSCO Shipbuilding' },
    { keywords: /mitsubishi/i, name: 'Mitsubishi Heavy Industries' },
    { keywords: /mhi/i, name: 'Mitsubishi Heavy Industries' },
    { keywords: /kawasaki/i, name: 'Kawasaki Heavy Industries' },
    { keywords: /nacks/i, name: 'NACKS Shipbuilding' },
    { keywords: /new times/i, name: 'New Times Shipbuilding' },
    { keywords: /stx/i, name: 'STX Offshore & Shipbuilding' },
    { keywords: /hanjin/i, name: 'Hanjin Heavy Industries' },
    { keywords: /guangzhou/i, name: 'Guangzhou Shipyard International' },
    { keywords: /philippine/i, name: 'Philippine Shipyard' },
    { keywords: /shin kurushima/i, name: 'Shin Kurushima Dockyard' },
    { keywords: /namura/i, name: 'Namura Shipbuilding' },
    { keywords: /jiangsu/i, name: 'Jiangsu Yangzijiang Shipbuilding' },
    { keywords: /sanoyas/i, name: 'Sanoyas Shipbuilding' },
  ];

  /** Known vessel type keyword → canonical term. */
  private static readonly VESSEL_TYPE_MAP: Array<{ keywords: RegExp; type: string }> = [
    { keywords: /bulk.carrier|bulker|obo/i, type: 'Bulk Carrier' },
    { keywords: /vlcc|ultra.large.crude/i, type: 'VLCC' },
    { keywords: /suezmax/i, type: 'Suezmax Tanker' },
    { keywords: /aframax/i, type: 'Aframax Tanker' },
    { keywords: /product.tanker|clean.tanker|lng.tanker/i, type: 'Product Tanker' },
    { keywords: /lng\b/i, type: 'LNG Carrier' },
    { keywords: /lpg\b/i, type: 'LPG Carrier' },
    { keywords: /chemical.tanker/i, type: 'Chemical Tanker' },
    { keywords: /crude.tanker|tanker/i, type: 'Tanker' },
    { keywords: /container|feeder/i, type: 'Container Ship' },
    { keywords: /ro.ro|roro|car.carrier/i, type: 'RoRo / Car Carrier' },
    { keywords: /general.cargo/i, type: 'General Cargo' },
    { keywords: /offshore/i, type: 'Offshore Vessel' },
    { keywords: /ferry|pax|passenger/i, type: 'Passenger / Ferry' },
  ];

  /**
   * Extracts vessel identity information from a list of uploaded file names.
   * Only reads file metadata (names) — no network calls.
   */
  /**
   * Extracts vessel identity information using backend OCR Staging extraction + file metadata.
   * Leverages real OCR text, verified vessel names, and classification tags.
   */
  /**
   * Extracts vessel identity information using backend OCR Staging extraction + file metadata.
   * Leverages real OCR text, verified vessel names, classification tags, and filename patterns.
   */
  public async _extractVesselInfoFromFiles(
    files: File[],
    contextVesselName?: string,
    uploadEntries?: VesselSuggestionUploadEntry[],
  ): Promise<{ primary: VesselSuggestion; all: VesselSuggestion[] }> {
    const names = files.map(f => f.name);
    const namesSet = new Set(names.map(n => n.trim().toLowerCase()));

    // 1. Fetch live OCR Staging items from backend
    let ocrItems: OcrStagingItem[] = [];
    try {
      const base = this._base();
      if (base) {
        const res = await fetch(`${base}/api/ocr/staging`);
        if (res.ok) {
          ocrItems = await res.json();
        }
      }
    } catch (e) {
      console.warn('[VesselDMS] Could not fetch OCR staging items for vessel suggestions:', e);
    }

    // 2. Filter OCR staging items relevant to the uploaded files.
    // Prefer drive_item_id matching first, then filename as fallback.
    const uploadDriveIds = new Set((uploadEntries || []).map(e => (e.drive_item_id || '').trim()).filter(Boolean));
    let matchingOcr = ocrItems.filter(item => {
      const itemDid = (item.drive_item_id || '').trim();
      if (uploadDriveIds.size > 0 && itemDid) {
        return uploadDriveIds.has(itemDid);
      }
      return (
        namesSet.size === 0 ||
        namesSet.has((item.filename || '').trim().toLowerCase()) ||
        names.some(n => (item.filename || '').toLowerCase().includes(n.toLowerCase()) || n.toLowerCase().includes((item.filename || '').toLowerCase()))
      );
    });
    if (matchingOcr.length === 0 && ocrItems.length > 0 && namesSet.size === 0) {
      matchingOcr = ocrItems;
    }

    // 3. Group OCR findings by detected vessel name
    const suggestionsByVessel: Map<string, VesselSuggestion> = new Map();

    for (const item of matchingOcr) {
      if (item.status === 'needs_review') continue;
      const tags = item.suggested_tags || {};
      const tagVessel = typeof tags.vessel === 'object' && tags.vessel !== null ? String(tags.vessel.value || '') : String(tags.vessel || '');
      const detectedVessel = (item.vessel_name || tagVessel || tags.vessel_name || '').trim();
      if (!detectedVessel || detectedVessel.toLowerCase() === 'common for all ships') continue;

      const imo = String(tags.imo || tags.imo_number || tags.imoNumber || '').trim();
      const shipyard = String(tags.shipyard || tags.shipyard_name || '').trim();
      const hull = String(tags.hull || tags.hull_number || tags.hullNumber || '').trim();
      const vType = String(tags.vessel_type || tags.vesselType || tags.category || 'Bulk Carrier').trim();

      const existing = suggestionsByVessel.get(detectedVessel.toLowerCase());
      if (existing) {
        if (!existing.imoNumber && imo) existing.imoNumber = imo;
        if (!existing.shipyard && shipyard) existing.shipyard = shipyard;
        if (!existing.hullNumber && hull) existing.hullNumber = hull;
        if (!existing.sourceFiles.includes(item.filename)) existing.sourceFiles.push(item.filename);
        if (item.confidence && item.confidence > (existing.ocrConfidence || 0)) {
          existing.ocrConfidence = item.confidence;
          existing.confidence = Math.max(existing.confidence, item.confidence);
        }
        if (item.matched_keywords && item.matched_keywords.length > 0) {
          existing.matchedKeywords = Array.from(new Set([...(existing.matchedKeywords || []), ...item.matched_keywords]));
        }
      } else {
        suggestionsByVessel.set(detectedVessel.toLowerCase(), {
          vesselName: detectedVessel,
          imoNumber: imo,
          shipyard: shipyard,
          hullNumber: hull,
          vesselType: vType || 'Bulk Carrier',
          matchedExisting: null,
          confidence: item.confidence || 0.95,
          sourceFiles: [item.filename],
          ocrVerified: true,
          ocrConfidence: item.confidence || 0.95,
          ocrTextPreview: item.ocr_text_preview || undefined,
          matchedKeywords: item.matched_keywords || [],
          suggestedTags: tags,
        });
      }
    }

    // 4. File name scanning for explicit vessel name prefixes (e.g. "GHANA EXPRESS - General Arrangement.pdf")
    const allText = names.join(' ');
    const imoMatch = allText.match(/\bIMO[.\- ]?(\d{7})\b/i) || allText.match(/\b(\d{7})\b/);
    const textImo = imoMatch ? imoMatch[1] : '';

    let textHull = '';
    const hullPatterns = [
      /\b(SS\d{3,6})\b/i,
      /\b(HN[ -]?\d{3,6})\b/i,
      /\b(YN[ -]?\d{3,6})\b/i,
      /\b(NB[ -]?\d{3,6})\b/i,
      /^([A-Z]{1,4}\d{3,6})/,
    ];
    for (const p of hullPatterns) {
      const m = allText.match(p) || (names[0] && names[0].match(p));
      if (m) { textHull = m[1].replace(/\s/g, '').toUpperCase(); break; }
    }

    let textShipyard = '';
    for (const entry of VesselEmail.SHIPYARD_MAP) {
      if (entry.keywords.test(allText)) { textShipyard = entry.name; break; }
    }

    let textType = '';
    for (const entry of VesselEmail.VESSEL_TYPE_MAP) {
      if (entry.keywords.test(allText)) { textType = entry.type; break; }
    }

    for (const name of names) {
      const cleanFn = name.replace(/\.[^/.]+$/, '');
      const dashMatch = cleanFn.match(/^([A-Z0-9\s]{3,30}?)\s*[-–—_]\s*(.+)$/i);
      if (dashMatch) {
        const candidate = dashMatch[1].trim();
        const candNorm = candidate.toLowerCase();
        const genericWords = new Set([
          'drawing', 'manual', 'certificate', 'plan', 'report', 'spec', 'file', 'doc', 'pdf',
          '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
          'empty', 'eng', 'test', 'result', 'results', 'shop', 'screenshot', 'temp', 'scan',
          'scanned', 'image', 'img', 'document', 'untitled', 'page', 'section', 'table', 'index'
        ]);
        if (!genericWords.has(candNorm) && candidate.length >= 3 && !/^\d+$/.test(candidate)) {
          const formattedName = candidate.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
          const existing = suggestionsByVessel.get(candNorm);
          if (existing) {
            if (!existing.sourceFiles.includes(name)) existing.sourceFiles.push(name);
          } else {
            suggestionsByVessel.set(candNorm, {
              vesselName: formattedName,
              imoNumber: textImo,
              shipyard: textShipyard,
              hullNumber: textHull,
              vesselType: textType || 'Bulk Carrier',
              matchedExisting: null,
              confidence: 0.95,
              sourceFiles: [name],
              ocrVerified: true,
              ocrConfidence: 0.95,
            });
          }
        }
      }
    }

    // If context vessel is passed, only use as fallback if no vessels were detected from OCR or filenames
    if (contextVesselName && contextVesselName !== 'Common for all ships' && contextVesselName !== 'Kaizen - Knowledge Bank') {
      if (suggestionsByVessel.size === 0) {
        suggestionsByVessel.set(contextVesselName.toLowerCase(), {
          vesselName: contextVesselName,
          imoNumber: textImo,
          shipyard: textShipyard,
          hullNumber: textHull,
          vesselType: textType || 'Bulk Carrier',
          matchedExisting: null,
          confidence: 0.80,
          sourceFiles: names,
        });
      }
    }

    // If no suggestions were found from OCR or filenames, create one from analysis
    if (suggestionsByVessel.size === 0) {
      const fallbackName = contextVesselName || (textHull ? `Vessel ${textHull}` : 'New Vessel');
      const fallbackSuggestion: VesselSuggestion = {
        vesselName: fallbackName,
        imoNumber: textImo,
        shipyard: textShipyard,
        hullNumber: textHull,
        vesselType: textType || 'Bulk Carrier',
        matchedExisting: null,
        confidence: 0,
        sourceFiles: names,
      };
      return { primary: fallbackSuggestion, all: [fallbackSuggestion] };
    }

    // Sort suggestions so the vessel with the most source files is first (primary)
    const allList = Array.from(suggestionsByVessel.values()).sort(
      (a, b) => (b.sourceFiles?.length || 0) - (a.sourceFiles?.length || 0)
    );
    const primary = allList[0];
    return { primary, all: allList };
  }

  /**
   * Tries to find an existing vessel in `this.state.vessels` that matches the
   * extracted IMO, hull number, or name. Returns the match + confidence score.
   */
  public _matchVesselInTermStore(suggestion: VesselSuggestion): { vessel: VesselRecord | null; confidence: number } {
    const { vessels } = this.state;
    if (!vessels.length) return { vessel: null, confidence: 0 };

    const ALIAS_MAP: Record<string, string[]> = {
      "Belle Lune": ["belle lune", "ss268", "ss 268", "ss-268", "ss_268", "hull ss268", "hull 268"],
      "Bow Fighter": ["bow fighter"],
      "Bow Fraternity": ["bow fraternity"],
      "Cameroun Express": ["cameroun express", "cameroon express"],
      "Cecilie F": ["cecilie f", "cecilie-f"],
      "Cote D Ivoire Express": ["cote d ivoire express", "cote d'ivoire express", "cote d' ivoire express", "côte d'ivoire express", "cote divoire express"],
      "Dutches Emerald": ["dutches emerald", "duchess emerald", "dutchess emerald"],
      "Ghana Express": ["ghana express"],
      "Lignum Grid": ["lignum grid"],
      "Lignum Mesh": ["lignum mesh"],
      "Lignum Web": ["lignum web"],
      "Maersk EI Banco": ["maersk ei banco", "maersk el banco", "ei banco", "el banco"],
      "Maersk EI Palomar": ["maersk ei palomar", "maersk el palomar", "ei palomar", "el palomar"],
      "Maersk Ferrato": ["maersk ferrato"],
      "Maersk Finisterre": ["maersk finisterre"],
      "Maersk Frio": ["maersk frio"],
      "Norse Evolution": ["norse evolution"],
      "Norse Ijmuiden": ["norse ijmuiden", "norse ymuiden"],
      "Norse New Haven": ["norse new haven"],
      "Peissy": ["peissy", "ss378", "ss 378", "ss-378", "ss_378", "hull ss378", "hull 378", "s.no.ss378"],
      "Potiniere": ["potiniere", "potinière"],
      "Senegal Express": ["senegal express"],
      "Snow Flake": ["snow flake", "snowflake"],
      "Snow Flower": ["snow flower", "snowflower"],
    };

    const sName = (suggestion.vesselName || '').toLowerCase().trim();
    const sHull = (suggestion.hullNumber || '').toLowerCase().trim();

    let bestVessel: VesselRecord | null = null;
    let bestScore = 0;

    for (const v of vessels) {
      let score = 0;
      const vNameLower = v.name.toLowerCase().trim();

      // Check Alias map
      const aliases = ALIAS_MAP[v.name] || [];
      if (aliases.some(a => (sName && (sName === a || sName.includes(a))) || (sHull && (sHull === a || sHull.includes(a))))) {
        score += 0.95;
      }

      // Direct name match (from OCR or folder context) → very high confidence
      if (suggestion.vesselName && vNameLower === sName) score += 0.90;
      // IMO exact match → very high confidence
      if (suggestion.imoNumber && v.imo && v.imo.trim() === suggestion.imoNumber.trim()) score += 0.80;
      // Hull number exact match
      if (suggestion.hullNumber && v.hull_number &&
          v.hull_number.replace(/\s/g, '').toUpperCase() === suggestion.hullNumber.replace(/\s/g, '').toUpperCase()) score += 0.60;
      // Vessel name contains hull number (e.g. "Belle Lune SS268")
      if (suggestion.hullNumber && v.name.toUpperCase().includes(suggestion.hullNumber.toUpperCase())) score += 0.40;
      // Hull number appears in vessel name
      if (suggestion.hullNumber && v.hull_number &&
          suggestion.hullNumber.toUpperCase().includes(v.hull_number.replace(/\s/g, '').toUpperCase())) score += 0.30;
      // Partial vessel name match
      if (suggestion.vesselName && suggestion.vesselName.length > 3 &&
          vNameLower.includes(sName.split(' ')[0])) score += 0.25;
      // Shipyard match
      if (suggestion.shipyard && v.shipyard &&
          v.shipyard.toLowerCase().includes(suggestion.shipyard.split(' ')[0].toLowerCase())) score += 0.10;
      if (score > bestScore) { bestScore = score; bestVessel = v; }
    }

    // Consider a match at ≥ 20% confidence
    if (bestScore >= 0.20) return { vessel: bestVessel, confidence: Math.min(bestScore, 1) };
    return { vessel: null, confidence: 0 };
  }

  private async _findUnidentifiedVesselFiles(files: File[], uploadEntries?: VesselSuggestionUploadEntry[]): Promise<string[]> {
    if (!files.length) return [];
    const base = this._base();
    if (!base) return [];

    const targetNames = new Set(files.map(f => (f.name || '').trim().toLowerCase()).filter(Boolean));
    if (!targetNames.size) return [];

    const sleep = (ms: number): Promise<void> => new Promise(resolve => window.setTimeout(resolve, ms));

    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const res = await fetch(`${base}/api/ocr/staging`);
        if (!res.ok) {
          if (attempt < 3) await sleep(700);
          continue;
        }

        const allItems: OcrStagingItem[] = await res.json();
        const entryDriveIds = new Set((uploadEntries || []).map(e => (e.drive_item_id || '').trim()).filter(Boolean));
        let matched = (allItems || []).filter(item => targetNames.has((item.filename || '').trim().toLowerCase()));
        if (entryDriveIds.size > 0) {
          matched = matched.filter(item => (item.drive_item_id || '').trim() ? entryDriveIds.has((item.drive_item_id || '').trim()) : true);
        }

        if (matched.length === 0) {
          if (attempt < 3) await sleep(700);
          continue;
        }

        const unresolved = matched
          .filter(item => {
            if (item.status === 'needs_review') return true;
            const tags = item.suggested_tags || {};
            const rawVessel = (tags as any).vessel;
            const tagVesselVal = rawVessel && typeof rawVessel === 'object'
              ? String(rawVessel.value || '').trim()
              : String(rawVessel || '').trim();
            const tagVesselConf = rawVessel && typeof rawVessel === 'object' && typeof rawVessel.confidence === 'number'
              ? rawVessel.confidence
              : 0;
            const detected = (item.vessel_name || tagVesselVal || '').trim();
            if (!detected) return true;
            if (/^(unknown|not identified|n\/a|na)$/i.test(detected)) return true;
            if (tagVesselConf > 0 && tagVesselConf < 0.40) return true;
            return false;
          })
          .map(item => item.filename)
          .filter(Boolean);

        return Array.from(new Set(unresolved));
      } catch {
        if (attempt < 3) await sleep(700);
      }
    }

    return [];
  }

  private async _analyzeUploadQueueRouting(
    files: File[],
    uploadEntries?: VesselSuggestionUploadEntry[]
  ): Promise<{ unidentifiedFiles: string[]; queueTab: 'staged' | 'unstaged'; hasMatchedItems: boolean }> {
    if (!files.length) return { unidentifiedFiles: [], queueTab: 'staged', hasMatchedItems: false };
    const base = this._base();
    if (!base) return { unidentifiedFiles: [], queueTab: 'staged', hasMatchedItems: false };

    const targetNames = new Set(files.map(f => (f.name || '').trim().toLowerCase()).filter(Boolean));
    if (!targetNames.size) return { unidentifiedFiles: [], queueTab: 'staged', hasMatchedItems: false };

    const sleep = (ms: number): Promise<void> => new Promise(resolve => window.setTimeout(resolve, ms));

    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const res = await fetch(`${base}/api/ocr/staging`);
        if (!res.ok) {
          if (attempt < 3) await sleep(700);
          continue;
        }

        const allItems: OcrStagingItem[] = await res.json();
        const entryDriveIds = new Set((uploadEntries || []).map(e => (e.drive_item_id || '').trim()).filter(Boolean));
        let matched = (allItems || []).filter(item => targetNames.has((item.filename || '').trim().toLowerCase()));
        if (entryDriveIds.size > 0) {
          matched = matched.filter(item => (item.drive_item_id || '').trim() ? entryDriveIds.has((item.drive_item_id || '').trim()) : true);
        }

        if (matched.length === 0) {
          if (attempt < 3) await sleep(700);
          continue;
        }

        const unresolved = matched
          .filter(item => {
            if (item.status === 'needs_review') return true;
            const tags = item.suggested_tags || {};
            const rawVessel = (tags as any).vessel;
            const tagVesselVal = rawVessel && typeof rawVessel === 'object'
              ? String(rawVessel.value || '').trim()
              : String(rawVessel || '').trim();
            const tagVesselConf = rawVessel && typeof rawVessel === 'object' && typeof rawVessel.confidence === 'number'
              ? rawVessel.confidence
              : 0;
            const detected = (item.vessel_name || tagVesselVal || '').trim();
            if (!detected) return true;
            if (/^(unknown|not identified|n\/a|na)$/i.test(detected)) return true;
            if (tagVesselConf > 0 && tagVesselConf < 0.40) return true;
            return false;
          })
          .map(item => item.filename)
          .filter(Boolean);

        const queueTab: 'staged' | 'unstaged' = matched.some(item => String(item.status || '').toLowerCase() === 'needs_review')
          ? 'unstaged'
          : 'staged';

        return {
          unidentifiedFiles: Array.from(new Set(unresolved)),
          queueTab,
          hasMatchedItems: true,
        };
      } catch {
        if (attempt < 3) await sleep(700);
      }
    }

    return { unidentifiedFiles: [], queueTab: 'staged', hasMatchedItems: false };
  }

  /**
   * Opens the Vessel Suggestions wizard. Called after file uploads or from sidebar.
   * @param files          The uploaded File objects to analyse.
   * @param vesselName     Optional vessel name from the upload folder context.
   */
  public async _openVesselSuggestions(
    files: File[],
    vesselName?: string,
    uploadEntries?: VesselSuggestionUploadEntry[],
  ): Promise<void> {
    // Disabled OCR staging review queue routing per user request: we do not want review & ocr

    const dialog: VesselSuggestionDialog = {
      open: true,
      step: 'extract',
      files: files || [],
      uploadEntries: uploadEntries || [],
      extracting: true,
      suggestion: null,
      editedSuggestion: null,
      allSuggestions: [],
      selectedSuggestionIndex: 0,
      creating: false,
      error: null,
    };
    this.setState({ vesselSuggestionDialog: dialog, ocrQueueTabHint: null, ocrUnidentifiedFiles: [] }, () => {
      void this._runVesselExtraction(files || [], vesselName, uploadEntries || []);
    });
  }

  private _getTagString(tags: Record<string, any> | null | undefined, key: string): string {
    if (!tags) return '';
    const raw = (tags as any)[key];
    if (raw && typeof raw === 'object') return String(raw.value || '').trim();
    return String(raw || '').trim();
  }

  private _pickLatestStagingItem(items: OcrStagingItem[]): OcrStagingItem | null {
    if (!items.length) return null;
    return items
      .slice()
      .sort((a, b) => {
        const ta = Date.parse(a.created_at || '') || 0;
        const tb = Date.parse(b.created_at || '') || 0;
        return tb - ta;
      })[0] || null;
  }

  private _normalizeVesselName(value: string): string {
    return (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  private _getSuggestionKey(s: VesselSuggestion): string {
    const imo = (s.imoNumber || '').trim();
    if (/^\d{7}$/.test(imo)) return `imo:${imo}`;

    const hull = (s.hullNumber || '').trim().replace(/\s/g, '').toUpperCase();
    if (hull) return `hull:${hull}`;

    const name = this._normalizeVesselName(s.vesselName || '');
    return `name:${name}`;
  }

  private _mergePendingSuggestions(existing: VesselSuggestion[], incoming: VesselSuggestion[]): VesselSuggestion[] {
    const merged: VesselSuggestion[] = [];
    const seen = new Set<string>();
    const ignored = this.state.ignoredVesselSuggestionKeys || new Set<string>();
    const sortedIncoming = [...incoming].sort((a, b) => (b.lastDetectedAt || 0) - (a.lastDetectedAt || 0));
    const sortedExisting = [...existing].sort((a, b) => (b.lastDetectedAt || 0) - (a.lastDetectedAt || 0));
    for (const suggestion of [...sortedIncoming, ...sortedExisting]) {
      if (!suggestion) continue;
      const key = this._getSuggestionKey(suggestion);
      if (!key || ignored.has(key) || seen.has(key)) continue;
      seen.add(key);
      merged.push(suggestion);
    }
    return merged.sort((a, b) => (b.lastDetectedAt || 0) - (a.lastDetectedAt || 0));
  }

  private _isKnownVesselSuggestionInState(suggestion: VesselSuggestion): boolean {
    const vessels = this.state.vessels || [];
    const norm = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const sName = norm(suggestion.vesselName || '');
    const sImo = (suggestion.imoNumber || '').trim();
    const sHull = (suggestion.hullNumber || '').trim().replace(/\s/g, '').toUpperCase();

    for (const v of vessels) {
      const vName = norm(v.name || '');
      const vImo = (v.imo || '').trim();
      const vHull = (v.hull_number || '').trim().replace(/\s/g, '').toUpperCase();
      if (sName && vName && sName === vName) return true;
      if (sImo && vImo && sImo === vImo) return true;
      if (sHull && vHull && sHull === vHull) return true;
    }
    return false;
  }

  public _syncPendingVesselSuggestionsFromStaging = async (
    stagingItems: OcrStagingItem[],
    autoOpen: boolean = false,
  ): Promise<void> => {
    if (!Array.isArray(stagingItems) || stagingItems.length === 0) return;

    const vesselConfidenceFloor = 0.60;
    const excludedNames = new Set(['common for all ships', 'common for all vessels', 'kaizen - knowledge bank']);
    const byKey = new Map<string, VesselSuggestion>();

    for (const item of stagingItems) {
      const status = String(item?.status || '').toLowerCase();
      if (!item || status === 'moved' || status === 'dismissed') continue;

      const tags = (item.suggested_tags || {}) as Record<string, any>;
      const rawVessel = tags.vessel;
      const vesselVal = rawVessel && typeof rawVessel === 'object'
        ? String(rawVessel.value || '').trim()
        : String(rawVessel || '').trim();
      const vesselConfidence = rawVessel && typeof rawVessel === 'object' && typeof rawVessel.confidence === 'number'
        ? Number(rawVessel.confidence)
        : (typeof item.confidence === 'number' ? Number(item.confidence) : 0);

      const detectedVessel = (item.vessel_name || vesselVal || '').trim();
      if (!detectedVessel) continue;
      if (excludedNames.has(detectedVessel.toLowerCase())) continue;
      if (vesselConfidence > 0 && vesselConfidence < vesselConfidenceFloor) continue;

      const candidate: VesselSuggestion = {
        vesselName: detectedVessel,
        imoNumber: this._getTagString(tags, 'imo') || this._getTagString(tags, 'imo_number') || this._getTagString(tags, 'imoNumber'),
        shipyard: this._getTagString(tags, 'shipyard') || this._getTagString(tags, 'shipyard_name'),
        hullNumber: this._getTagString(tags, 'hull') || this._getTagString(tags, 'hull_number') || this._getTagString(tags, 'hullNumber'),
        vesselType: this._getTagString(tags, 'vessel_type') || this._getTagString(tags, 'vesselType') || 'Bulk Carrier',
        matchedExisting: null,
        confidence: typeof item.confidence === 'number' ? item.confidence : vesselConfidence,
        sourceFiles: [item.filename].filter(Boolean),
        ocrVerified: true,
        ocrConfidence: vesselConfidence,
        ocrTextPreview: item.ocr_text_preview || undefined,
        matchedKeywords: Array.isArray(item.matched_keywords) ? item.matched_keywords : [],
        suggestedTags: tags,
        lastDetectedAt: Date.parse(item.updated_at || item.created_at || '') || Date.now(),
      };

      if (this._isKnownVesselSuggestionInState(candidate)) continue;

      const key = this._getSuggestionKey(candidate);
      if (!key) continue;
      const existing = byKey.get(key);
      if (!existing) {
        byKey.set(key, candidate);
      } else {
        const fileSet = new Set([...(existing.sourceFiles || []), ...(candidate.sourceFiles || [])]);
        existing.sourceFiles = Array.from(fileSet);
        existing.confidence = Math.max(existing.confidence || 0, candidate.confidence || 0);
        existing.ocrConfidence = Math.max(existing.ocrConfidence || 0, candidate.ocrConfidence || 0);
        existing.lastDetectedAt = Math.max(existing.lastDetectedAt || 0, candidate.lastDetectedAt || 0);
        if (!existing.imoNumber && candidate.imoNumber) existing.imoNumber = candidate.imoNumber;
        if (!existing.hullNumber && candidate.hullNumber) existing.hullNumber = candidate.hullNumber;
        if (!existing.shipyard && candidate.shipyard) existing.shipyard = candidate.shipyard;
      }
    }

    const incoming = Array.from(byKey.values());
    if (incoming.length === 0) return;

    this.setState(prev => {
      const ignored = prev.ignoredVesselSuggestionKeys || new Set<string>();
      const filteredIncoming = incoming.filter(s => !ignored.has(this._getSuggestionKey(s)));
      if (filteredIncoming.length === 0) return null;

      const prevKeys = new Set((prev.pendingVesselSuggestions || []).map(s => this._getSuggestionKey(s)));
      const mergedPending = this._mergePendingSuggestions(prev.pendingVesselSuggestions || [], filteredIncoming);
      const newlyAdded = filteredIncoming.filter(s => !prevKeys.has(this._getSuggestionKey(s)));

      if (!autoOpen || newlyAdded.length === 0 || prev.vesselSuggestionDialog) {
        return { pendingVesselSuggestions: mergedPending } as any;
      }

      const primary = newlyAdded.sort((a, b) => (b.lastDetectedAt || 0) - (a.lastDetectedAt || 0))[0];
      const sourceFiles = Array.from(new Set(
        mergedPending.flatMap(s => (s.sourceFiles || []).map(name => (name || '').trim()).filter(Boolean))
      )).map(name => new File([], name));

      return {
        pendingVesselSuggestions: mergedPending,
        vesselSuggestionDialog: {
          open: true,
          step: 'review',
          files: sourceFiles,
          uploadEntries: [],
          extracting: false,
          suggestion: primary,
          editedSuggestion: { ...primary },
          allSuggestions: mergedPending,
          selectedSuggestionIndex: 0,
          creating: false,
          error: null,
        },
      } as any;
    });
  };

  public _ignoreCurrentVesselSuggestion = (): void => {
    const dlg = this.state.vesselSuggestionDialog;
    const current = dlg?.suggestion;
    if (!dlg || !current) return;

    const key = this._getSuggestionKey(current);
    const ignored = new Set(this.state.ignoredVesselSuggestionKeys || new Set<string>());
    if (key) ignored.add(key);
    this._persistIgnoredVesselSuggestionKeys(ignored);

    const remaining = (this.state.pendingVesselSuggestions || []).filter(s => this._getSuggestionKey(s) !== key);
    if (remaining.length === 0) {
      this.setState({
        ignoredVesselSuggestionKeys: ignored,
        pendingVesselSuggestions: [],
        vesselSuggestionDialog: null,
      });
      return;
    }

    const next = remaining[0];
    const nextFiles = Array.from(new Set(
      remaining.flatMap(s => (s.sourceFiles || []).map(name => (name || '').trim()).filter(Boolean))
    )).map(name => new File([], name));

    this.setState({
      ignoredVesselSuggestionKeys: ignored,
      pendingVesselSuggestions: remaining,
      vesselSuggestionDialog: {
        ...dlg,
        open: true,
        step: 'review',
        files: nextFiles,
        extracting: false,
        suggestion: next,
        editedSuggestion: { ...next },
        allSuggestions: remaining,
        selectedSuggestionIndex: 0,
        creating: false,
        error: null,
      },
    });
  };

  public _clearIgnoredVesselSuggestionKeys = (): void => {
    const empty = new Set<string>();
    this._persistIgnoredVesselSuggestionKeys(empty);
    const restoredPending = this._readPendingVesselSuggestions();
    this.setState({
      ignoredVesselSuggestionKeys: empty,
      pendingVesselSuggestions: restoredPending,
    });
  };

  public async _applyVesselSuggestionUseExisting(): Promise<void> {
    const dlg = this.state.vesselSuggestionDialog;
    const matched = dlg?.suggestion?.matchedExisting ?? null;
    const matchedVesselName = (matched?.name || '').trim();
    if (!dlg || !matchedVesselName) return;

    this.setState(prev => ({
      vesselSuggestionDialog: prev.vesselSuggestionDialog
        ? { ...prev.vesselSuggestionDialog, creating: true, error: null }
        : null,
    }));

    try {
      const base = this._base();
      if (!base) throw new Error('API base URL is not configured.');

      const listRes = await fetch(`${base}/api/ocr/staging`, { headers: this._headers() });
      if (!listRes.ok) throw new Error(`Could not load OCR staging queue (status ${listRes.status}).`);
      const allItems: OcrStagingItem[] = await listRes.json();

      const allowedStatuses = new Set(['tag_suggested', 'needs_review']);
      const activeItems = (allItems || []).filter(i => allowedStatuses.has(String(i.status || '').toLowerCase()));

      const sourceFiles = (dlg.files || []).map(f => (f.name || '').trim()).filter(Boolean);
      const sourceFileSet = new Set(sourceFiles.map(n => n.toLowerCase()));
      const uploadEntries = (dlg.uploadEntries || []).filter(e => (e.filename || '').trim());

      const matchedRows = new Map<number, OcrStagingItem>();

      // Primary match: drive_item_id + allowed status (most reliable, no stale filename collisions).
      for (const entry of uploadEntries) {
        const did = (entry.drive_item_id || '').trim();
        if (!did) continue;
        const byDrive = activeItems.filter(i => (i.drive_item_id || '').trim() === did);
        const pick = this._pickLatestStagingItem(byDrive);
        if (pick) matchedRows.set(pick.id, pick);
      }

      // Fallback match: filename + allowed status + most recent created_at.
      if (matchedRows.size === 0) {
        const names = uploadEntries.length > 0
          ? uploadEntries.map(e => (e.filename || '').trim().toLowerCase()).filter(Boolean)
          : Array.from(sourceFileSet.values());
        const uniqueNames = Array.from(new Set(names));
        for (const nameLower of uniqueNames) {
          const byName = activeItems.filter(i => (i.filename || '').trim().toLowerCase() === nameLower);
          const pick = this._pickLatestStagingItem(byName);
          if (pick) matchedRows.set(pick.id, pick);
        }
      }

      // Last-chance fallback from suggestion sourceFiles.
      if (matchedRows.size === 0 && Array.isArray(dlg.suggestion?.sourceFiles)) {
        const suggestionNames = Array.from(new Set((dlg.suggestion?.sourceFiles || []).map(n => (n || '').trim().toLowerCase()).filter(Boolean)));
        for (const nameLower of suggestionNames) {
          const byName = activeItems.filter(i => (i.filename || '').trim().toLowerCase() === nameLower);
          const pick = this._pickLatestStagingItem(byName);
          if (pick) matchedRows.set(pick.id, pick);
        }
      }

      const targets = Array.from(matchedRows.values());
      if (targets.length === 0) {
        throw new Error('No active OCR staging item found for this upload. Please review the file in Templates & OCR queue.');
      }

      const suggestionTags = (dlg.suggestion?.suggestedTags || {}) as Record<string, any>;
      const preferredGroup = this._getTagString(suggestionTags, 'group');
      const preferredCategory = this._getTagString(suggestionTags, 'category');
      const preferredSub = this._getTagString(suggestionTags, 'sub_category') || this._getTagString(suggestionTags, 'subcategory');

      const moveResults: Array<{ id: number; filename: string; target_path?: string }> = [];
      const moveErrors: string[] = [];
      let crossVesselReroutes = 0;

      for (const item of targets) {
        try {
          const currentTags = (item.suggested_tags || {}) as Record<string, any>;
          const sourceVessel = (item.vessel_name || this._getTagString(currentTags, 'vessel') || '').trim();
          if (sourceVessel && sourceVessel.toLowerCase() !== matchedVesselName.toLowerCase()) {
            crossVesselReroutes += 1;
          }

          const mergedTags: Record<string, any> = {
            ...currentTags,
            // Explicitly override vessel target so files are relocated to the matched vessel branch.
            vessel: matchedVesselName,
          };

          if (preferredGroup) mergedTags.group = preferredGroup;
          if (preferredCategory) mergedTags.category = preferredCategory;
          if (preferredSub) {
            mergedTags.sub_category = preferredSub;
            mergedTags.subcategory = preferredSub;
          }

          const patchRes = await fetch(`${base}/api/ocr/staging/${item.id}/tags`, {
            method: 'PATCH',
            headers: this._headers(),
            body: JSON.stringify({
              category_id: item.category_id,
              suggested_tags: mergedTags,
            }),
          });
          if (!patchRes.ok) {
            const patchErr = await patchRes.json().catch(() => ({}));
            throw new Error(patchErr?.detail || `Tag update failed (status ${patchRes.status})`);
          }

          const moveRes = await fetch(`${base}/api/ocr/staging/${item.id}/move`, {
            method: 'POST',
            headers: this._headers(),
          });
          if (!moveRes.ok) {
            const moveErr = await moveRes.json().catch(() => ({}));
            throw new Error(moveErr?.detail || `Move failed (status ${moveRes.status})`);
          }
          const moved = await moveRes.json().catch(() => ({}));
          moveResults.push({ id: item.id, filename: item.filename, target_path: moved?.target_path });
        } catch (err: any) {
          moveErrors.push(`${item.filename}: ${err?.message || 'Move failed'}`);
        }
      }

      if (moveResults.length === 0) {
        throw new Error(moveErrors[0] || 'No files were moved.');
      }

      const mergedVessels = new Set<string>([matchedVesselName]);
      const dialogSourceVessel = (dlg.uploadEntries || []).map(e => (e.source_vessel_name || '').trim()).filter(Boolean);
      dialogSourceVessel.forEach(v => mergedVessels.add(v));

      this.setState(prev => ({
        vesselSuggestionDialog: prev.vesselSuggestionDialog
          ? {
              ...prev.vesselSuggestionDialog,
              creating: false,
              step: 'done',
              editedSuggestion: {
                ...(prev.vesselSuggestionDialog.editedSuggestion || prev.vesselSuggestionDialog.suggestion || {
                  vesselName: matchedVesselName,
                  imoNumber: '',
                  shipyard: '',
                  hullNumber: '',
                  vesselType: 'Bulk Carrier',
                  matchedExisting: matched,
                  confidence: 1,
                  sourceFiles: sourceFiles,
                }),
                vesselName: matchedVesselName,
                matchedExisting: matched,
              },
              error: moveErrors.length > 0 ? `Moved ${moveResults.length} file(s), ${moveErrors.length} failed. ${moveErrors[0]}` : null,
            }
          : null,
        pendingVesselSuggestions: prev.pendingVesselSuggestions.filter(s => this._getSuggestionKey(s) !== this._getSuggestionKey(dlg.suggestion || {
          vesselName: matchedVesselName,
          imoNumber: '',
          shipyard: '',
          hullNumber: '',
          vesselType: '',
          matchedExisting: null,
          confidence: 0,
          sourceFiles: [],
        })),
        docUploadMsg: crossVesselReroutes > 0
          ? `Moved ${moveResults.length} file(s) to matched vessel "${matchedVesselName}". Corrected ${crossVesselReroutes} cross-vessel placement(s).`
          : `Moved ${moveResults.length} file(s) to matched vessel "${matchedVesselName}".`,
      }));

      void this._syncScheduler?.triggerNow().catch(() => undefined);
      void this._mergeLiveSharePointFiles(Array.from(mergedVessels)).catch(() => undefined);
    } catch (e: any) {
      this.setState(prev => ({
        vesselSuggestionDialog: prev.vesselSuggestionDialog
          ? { ...prev.vesselSuggestionDialog, creating: false, error: e?.message || 'Failed to route files to matched vessel.' }
          : null,
      }));
    }
  }

  /**
   * Thoroughly verifies whether a vessel exists by checking:
   * 1. Term Store & State vessels (`this.state.vessels`)
   * 2. Live Backend DB vessels (`/api/vessels` - catches any recently manually created vessels)
   * 3. Manually created vessels in SharePoint folder tree / DMS rows (`this.state.rows`)
   * 4. Normalized Name, IMO Number, or Hull Number matching
   */
  public async _isExistingVessel(
    suggestion: VesselSuggestion
  ): Promise<{ isExisting: boolean; matchedVessel: VesselRecord | null; reason?: string }> {
    const rawName = (suggestion.vesselName || '').trim();
    const rawImo = (suggestion.imoNumber || '').trim();
    const rawHull = (suggestion.hullNumber || '').trim().replace(/\s/g, '').toUpperCase();

    // Collect all known vessels from state
    const allVessels: VesselRecord[] = [...this.state.vessels];

    // Also fetch live from backend API to catch any freshly manually created vessels
    try {
      const base = this._base();
      if (base) {
        const res = await fetch(`${base}/api/vessels`);
        if (res.ok) {
          const liveList: VesselRecord[] = await res.json();
          if (Array.isArray(liveList) && liveList.length > 0) {
            const seenIds = new Set(allVessels.map(v => String(v.id || v.name).toLowerCase()));
            for (const lv of liveList) {
              if (!seenIds.has(String(lv.id || lv.name).toLowerCase())) {
                allVessels.push(lv);
                seenIds.add(String(lv.id || lv.name).toLowerCase());
              }
            }
          }
        }
      }
    } catch (e) {
      console.warn('[VesselDMS] _isExistingVessel live fetch error:', e);
    }

    const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
    const normTarget = norm(rawName);

    // 1. Check by Name (Exact & Normalized match)
    if (rawName && normTarget) {
      for (const v of allVessels) {
        if (!v.name) continue;
        const normV = norm(v.name);
        if (normV === normTarget) {
          return { isExisting: true, matchedVessel: v, reason: `Matched vessel name: ${v.name}` };
        }
      }
    }

    // 2. Check by IMO number
    if (rawImo && /^\d{7}$/.test(rawImo)) {
      for (const v of allVessels) {
        if (v.imo && v.imo.trim() === rawImo) {
          return { isExisting: true, matchedVessel: v, reason: `Matched IMO number ${rawImo} of vessel ${v.name}` };
        }
      }
    }

    // 3. Check by Hull number
    if (rawHull && rawHull.length >= 3) {
      for (const v of allVessels) {
        if (v.hull_number && v.hull_number.trim().replace(/\s/g, '').toUpperCase() === rawHull) {
          return { isExisting: true, matchedVessel: v, reason: `Matched Hull number ${rawHull} of vessel ${v.name}` };
        }
      }
    }

    return { isExisting: false, matchedVessel: null };
  }

  /** Runs the async extraction + match pipeline and updates dialog state. */
  private async _runVesselExtraction(
    files: File[],
    contextVesselName?: string,
    uploadEntries?: VesselSuggestionUploadEntry[],
  ): Promise<void> {
    try {
      // Step 1: extract from backend OCR staging + file metadata + filename patterns
      const { primary, all } = await this._extractVesselInfoFromFiles(files, contextVesselName, uploadEntries || []);
      const inferredDetectedAt = Math.max(0, ...((uploadEntries || []).map(e => Number(e.uploaded_at) || 0))) || Date.now();
      const allWithMeta = all.map(s => ({ ...s, lastDetectedAt: s.lastDetectedAt || inferredDetectedAt }));
      const primaryWithMeta = { ...primary, lastDetectedAt: primary.lastDetectedAt || inferredDetectedAt };
      await new Promise(r => setTimeout(r, 600));

      this.setState(prev => ({
        vesselSuggestionDialog: prev.vesselSuggestionDialog
          ? {
              ...prev.vesselSuggestionDialog,
              step: 'match',
              extracting: false,
              suggestion: primaryWithMeta,
              editedSuggestion: { ...primaryWithMeta },
              allSuggestions: allWithMeta,
              selectedSuggestionIndex: 0,
            }
          : null,
      }));

      // Step 2: match all against Term Store + Manually Created vessels + Live DB
      await new Promise(r => setTimeout(r, 500));
      const relevantSuggestions: VesselSuggestion[] = [];

      for (const s of allWithMeta) {
        const check = await this._isExistingVessel(s);
        const termStoreMatch = this._matchVesselInTermStore(s);
        const isExisting = check.isExisting || (termStoreMatch.vessel !== null && termStoreMatch.confidence >= 0.40);
        const matchedRecord = check.matchedVessel || termStoreMatch.vessel;

        // If the user uploaded inside an existing vessel folder (contextVesselName),
        // check whether the document belongs to the SAME vessel or a DIFFERENT/NEW vessel:
        const isSameAsCurrentFolder = contextVesselName && s.vesselName &&
          (s.vesselName.toLowerCase() === contextVesselName.toLowerCase() ||
           (matchedRecord?.name && matchedRecord.name.toLowerCase() === contextVesselName.toLowerCase()));

        if (!isSameAsCurrentFolder) {
          if (!isExisting) {
            // New uncreated vessel detected via OCR / filenames -> ALWAYS SUGGEST TO CREATE
            relevantSuggestions.push({
              ...s,
              matchedExisting: null,
              confidence: s.confidence,
            });
          } else if (contextVesselName && s.vesselName.toLowerCase() !== contextVesselName.toLowerCase()) {
            // Belongs to another existing vessel -> Suggest to route/link to that vessel
            relevantSuggestions.push({
              ...s,
              matchedExisting: matchedRecord,
              confidence: check.isExisting ? 1.0 : s.confidence,
            });
          }
        }
      }

      // If triggered from upload (files.length > 0) and NO new/different vessels were detected:
      if (files.length > 0 && relevantSuggestions.length === 0) {
        this.setState({
          vesselSuggestionDialog: null,
          docUploadMsg: 'Files successfully uploaded. All documents match the current vessel structure.',
        });
        return;
      }

      // If opened from sidebar and no new vessels found, provide a clean blank new vessel form
      const ignoredKeys = this.state.ignoredVesselSuggestionKeys || new Set<string>();
      const activeSuggestions = relevantSuggestions.filter(s => !ignoredKeys.has(this._getSuggestionKey(s)));
      const primarySuggestion = activeSuggestions.length > 0 ? activeSuggestions[0] : (files.length === 0 ? {
        vesselName: '',
        imoNumber: '',
        shipyard: '',
        hullNumber: '',
        vesselType: 'Bulk Carrier',
        matchedExisting: null,
        confidence: 0,
        sourceFiles: [],
        lastDetectedAt: Date.now(),
      } : null);

      if (!primarySuggestion && files.length > 0) {
        this.setState({ vesselSuggestionDialog: null });
        return;
      }

      this.setState(prev => ({
        pendingVesselSuggestions: this._mergePendingSuggestions(
          prev.pendingVesselSuggestions,
          activeSuggestions.filter(s => !s.matchedExisting)
        ),
        vesselSuggestionDialog: prev.vesselSuggestionDialog
          ? {
              ...prev.vesselSuggestionDialog,
              step: 'review',
              suggestion: primarySuggestion,
              allSuggestions: activeSuggestions,
              selectedSuggestionIndex: 0,
              editedSuggestion: primarySuggestion ? {
                ...primarySuggestion,
                vesselName: primarySuggestion.vesselName,
                imoNumber: primarySuggestion.imoNumber,
                shipyard: primarySuggestion.shipyard,
                hullNumber: primarySuggestion.hullNumber,
                vesselType: primarySuggestion.vesselType || 'Bulk Carrier',
              } : null,
            }
          : null,
      }));
    } catch (e: any) {
      this.setState(prev => ({
        vesselSuggestionDialog: prev.vesselSuggestionDialog
          ? { ...prev.vesselSuggestionDialog, extracting: false, error: e?.message || 'Extraction failed' }
          : null,
      }));
    }
  }

  /**
   * Opens the Suggested Vessels modal directly from the sidebar.
   */
  public _openSuggestedVesselsFromSidebar = (): void => {
    const pending = [...(this.state.pendingVesselSuggestions || [])].sort((a, b) => (b.lastDetectedAt || 0) - (a.lastDetectedAt || 0));
    if (pending.length > 0) {
      const sourceFileNames = Array.from(new Set(
        pending.flatMap(s => (s.sourceFiles || []).map(name => (name || '').trim()).filter(Boolean))
      ));
      const sourceFiles = sourceFileNames.map(name => new File([], name));
      const primary = pending[0];
      this.setState({
        vesselSuggestionDialog: {
          open: true,
          step: 'review',
          files: sourceFiles,
          uploadEntries: [],
          extracting: false,
          suggestion: primary,
          editedSuggestion: { ...primary },
          allSuggestions: pending,
          selectedSuggestionIndex: 0,
          creating: false,
          error: null,
        },
      });
      return;
    }

    void this._openVesselSuggestions([]);
  };


  /**
   * Creates a new vessel from the edited suggestion fields and provisions its folders.
   */
  public async _applyVesselSuggestionCreate(): Promise<void> {
    const { vesselSuggestionDialog } = this.state;
    if (!vesselSuggestionDialog?.editedSuggestion) return;
    const s = vesselSuggestionDialog.editedSuggestion;
    this.setState(prev => ({
      vesselSuggestionDialog: prev.vesselSuggestionDialog
        ? { ...prev.vesselSuggestionDialog, creating: true, error: null }
        : null,
    }));
    try {
      const newVessel: VesselRecord = {
        id: `v_${Date.now()}`,
        name: s.vesselName.trim(),
        imo: s.imoNumber.trim(),
        shipyard: s.shipyard.trim() || undefined,
        hull_number: s.hullNumber.trim() || undefined,
        vessel_type: s.vesselType || undefined,
        status: 'Active',
        image_url: pickRandomVesselImage(s.vesselType),
      };
      // POST to backend
      const res = await fetch(`${this._base()}/api/vessels`, {
        method: 'POST', headers: this._headers(),
        body: JSON.stringify({
          name: newVessel.name, imo: newVessel.imo,
          shipyard: newVessel.shipyard || null, hull_number: newVessel.hull_number || null,
          vessel_type: newVessel.vessel_type || null,
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (res.ok || res.status === 202) {
        if (data.id || data.result?.id) newVessel.id = data.id || data.result.id;
      }
      // Add to state + provision folders
      this.setState(prev => ({
        vessels: [newVessel, ...prev.vessels.filter(v => v.name.toLowerCase() !== newVessel.name.toLowerCase())],
        // Auto-navigate to Vessel Management after create from suggestions.
        view: 'vessels',
        selectedVessel: newVessel,
        vesselActionPicker: null,
        modal: 'none',
        vesselSuggestionDialog: null,
        pendingVesselSuggestions: prev.pendingVesselSuggestions.filter(sug => this._normalizeVesselName(sug.vesselName || '') !== this._normalizeVesselName(newVessel.name || '')),
      }));
      // Vessel creation must remain record-only. Do not create the default
      // Technical & Crewing / DMS folder template automatically after creation.
    } catch (e: any) {
      this.setState(prev => ({
        vesselSuggestionDialog: prev.vesselSuggestionDialog
          ? { ...prev.vesselSuggestionDialog, creating: false, error: e?.message || 'Failed to create vessel.' }
          : null,
      }));
    }
  }

  // ── Render Sidebar Navigation ─────────────────────────────────────────────


  public _renderSidebar(): React.ReactElement {
    return renderSidebar(this);
  }

  public _renderLayout(content: React.ReactElement): React.ReactElement {
    return renderLayout(this, content);
  }

  public _renderDocPreviewDrawer(): React.ReactElement | null {
    return renderDocPreviewDrawer(this);
  }

  public _renderDashboard(): React.ReactElement {
    return renderDashboard(this);
  }

  public _renderDocumentsPage(): React.ReactElement {
    return (
      <PageErrorBoundary pageName="Documents">
        <DocumentsPageWrapper host={this} />
      </PageErrorBoundary>
    );
  }

  public _renderVesselsPage(): React.ReactElement {
    return renderVesselsPage(this);
  }

  public _renderUsersPage(): React.ReactElement {
    return renderUsersPage(this);
  }

  /** Re-fetches the User Management list after a role/permission change so
   * the table reflects the just-saved state instead of the stale row the
   * edit started from. Mirrors the fetch in _goToView's 'users' branch. */
  private async _reloadUsersList(): Promise<void> {
    try {
      const data = await this._fetchJson(`${this._base()}/api/users`);
      const asRole = (value: any): UserItem['role'] => (String(value || '').toLowerCase() === 'admin' ? 'Admin' : 'User');
      const asStatus = (value: any): UserItem['status'] => (String(value || '').toLowerCase() === 'active' ? 'Active' : 'Inactive');
      const asPermissions = (value: any): UserItem['permissions'] => Array.isArray(value)
        ? value.map((p: any) => ({
          site_key: String(p?.site_key || ''),
          can_view: Boolean(p?.can_view),
          can_upload: Boolean(p?.can_upload),
          can_tag_on_upload: Boolean(p?.can_tag_on_upload),
        })).filter(p => p.site_key)
        : [];
      const users: UserItem[] = Array.isArray(data)
        ? data.map((row: any, idx: number) => {
          const email = String(row?.email || '').trim().toLowerCase();
          const fallbackName = email ? email.split('@')[0] : `user-${idx + 1}`;
          return {
            id: String(row?.id || email || `user-${idx + 1}`),
            name: String(row?.name || fallbackName),
            email,
            role: asRole(row?.role),
            status: asStatus(row?.status),
            lastLogin: String(row?.lastLogin || 'Never'),
            permissions: asPermissions(row?.permissions),
          };
        })
        : [];
      this.setState({ usersList: users });
    } catch {
      // Leave the existing list in place — a failed refresh shouldn't wipe
      // the table the admin is looking at.
    }
  }

  /** Admin-only: PATCH /api/admin/users/{email}/role (backend/app/main.py).
   * Requires the caller's own session to already be Admin — the endpoint
   * enforces that server-side via require_admin_session and 403s otherwise. */
  public async _updateUserRole(email: string, role: 'Admin' | 'User'): Promise<void> {
    this.setState({ usersPermissionsBusy: true });
    try {
      await this._fetchJson(`${this._base()}/api/admin/users/${encodeURIComponent(email)}/role`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ role }),
      });
      await this._reloadUsersList();
    } catch (err: any) {
      alert(err?.message === 'SESSION_EXPIRED' ? 'Your session expired — please sign in again.' : `Could not update role: ${err?.message || err}`);
    } finally {
      this.setState({ usersPermissionsBusy: false });
    }
  }

  /** Admin-only: PUT /api/admin/users/{email}/permissions (backend/app/main.py).
   * Replaces that user's entire site-permission set — the endpoint deletes
   * all existing UserSitePermission rows for the user and re-inserts the
   * given list, so `permissions` must be the full desired set, not a diff.
   * The backend also re-enforces the can_upload⊆can_view and
   * can_tag_on_upload⊆can_upload dependency server-side regardless of what
   * is sent here. */
  public async _updateUserSitePermissions(email: string, permissions: UserSitePermissionItem[]): Promise<void> {
    this.setState({ usersPermissionsBusy: true });
    try {
      await this._fetchJson(`${this._base()}/api/admin/users/${encodeURIComponent(email)}/permissions`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ permissions }),
      });
      await this._reloadUsersList();
      this.setState({ usersExpandedEmail: null });
    } catch (err: any) {
      alert(err?.message === 'SESSION_EXPIRED' ? 'Your session expired — please sign in again.' : `Could not update permissions: ${err?.message || err}`);
    } finally {
      this.setState({ usersPermissionsBusy: false });
    }
  }

  public _renderSettingsPage(): React.ReactElement {
    return renderSettingsPage(this);
  }

  public _renderProfilePage(): React.ReactElement {
    return renderProfilePage(this);
  }

  public _renderBentoEmailDashboardPage(): React.ReactElement {
    return renderBentoEmailDashboardPage(this);
  }


  // ── Bento Compose Modal ─────────────────────────────────────────────────

  // Build auto subject from vessel + tag ([DataSource:TAG] Vessel Name format)
  public _buildAutoSubject(vesselName: string, tag: string, customSubjectText: string = ''): string {
    const tagKey = tag || 'mail';
    const vName = (vesselName || '').trim();
    const text = (customSubjectText || '').trim();
    if (vName && text && text !== 'Subject') {
      return `[DataSource:${tagKey}] ${vName} / ${text}`;
    } else if (vName) {
      return `[DataSource:${tagKey}] ${vName}`;
    }
    return text && text !== 'Subject' ? `[DataSource:${tagKey}] ${text}` : `[DataSource:${tagKey}]`;
  }

  // Collect approved file names from rows & documents list for the selected vessel
  public _getApprovedFilesForVesselTag(vesselName: string, tag: string): string[] {
    const { rows, uploadedFilesByFolder, documentsList, bentoApprovedFiles } = this.state;
    const names: string[] = [];
    if (!vesselName || !vesselName.trim()) return names;

    const vLower = vesselName.trim().toLowerCase();

    // 1. From directly fetched approved files (most reliable)
    if (bentoApprovedFiles && bentoApprovedFiles[vLower]) {
      bentoApprovedFiles[vLower].forEach((f: string) => {
        if (names.indexOf(f) === -1) names.push(f);
      });
    }

    // 2. From flat rows
    rows.forEach(r => {
      if (r.vesselName.trim().toLowerCase() !== vLower) return;
      if (r.fileName && !r.filePending && names.indexOf(r.fileName) === -1) {
        names.push(r.fileName);
      }
      const uploads = uploadedFilesByFolder[r.groupKey] || [];
      (uploads || []).filter((f: any) => !f.pending).forEach((f: any) => {
        if (names.indexOf(f.name) === -1) names.push(f.name);
      });
    });

    // 3. From documents list
    (documentsList || []).forEach(doc => {
      if (doc.vessel && doc.vessel.trim().toLowerCase() === vLower && !doc.isFolder) {
        if (names.indexOf(doc.name) === -1) names.push(doc.name);
      }
    });

    return names;
  }

  // Fetch all approved files for a vessel directly from the backend
  public async _fetchApprovedFilesForVessel(vesselName: string): Promise<void> {
    if (!vesselName) return;
    const vLower = vesselName.trim().toLowerCase();

    try {
      const names: string[] = [];

      // 1. Trigger on-demand loading of files for this vessel if not already loaded
      if (!this._filesLoadedForVessels.has(vesselName)) {
        this._filesLoadedForVessels.add(vesselName);
        await this._mergeLiveSharePointFiles([vesselName]).catch(() => undefined);
      }

      // Approvals module removed — uploads are filed to SharePoint immediately
      // for every user, so files are already covered by step 1 (live merge)
      // and step 3 (rows/uploadedFilesByFolder) below.
      const fileIdUpdates: Record<string, string> = {};

      // 3. Collect from updated rows & uploadedFilesByFolder state after loading
      const { rows, uploadedFilesByFolder } = this.state;
      rows.forEach(r => {
        if (r.vesselName.trim().toLowerCase() !== vLower) return;
        if (r.fileName && !r.filePending && names.indexOf(r.fileName) === -1) {
          names.push(r.fileName);
        }
        (uploadedFilesByFolder[r.groupKey] || []).filter((f: any) => !f.pending).forEach((f: any) => {
          if (names.indexOf(f.name) === -1) names.push(f.name);
        });
      });

      // Always update state to populate dropdown
      this.setState(prev => ({
        bentoApprovedFiles: {
          ...(prev.bentoApprovedFiles || {}),
          [vLower]: names,
        },
        bentoApprovedFileIds: {
          ...(prev.bentoApprovedFileIds || {}),
          ...fileIdUpdates,
        },
      }));
    } catch {
      this.setState(prev => ({
        bentoApprovedFiles: {
          ...(prev.bentoApprovedFiles || {}),
          [vLower]: [],
        },
      }));
    }
  }

  public _getFileIdForAttachment(vesselName: string, fileName: string): string | null {
    const { rows, bentoApprovedFileIds } = this.state;
    const vLower = vesselName.trim().toLowerCase();
    for (const r of rows) {
      if (r.vesselName.trim().toLowerCase() === vLower && r.fileName === fileName && r.fileId) {
        return r.fileId;
      }
    }
    const approvedId = (bentoApprovedFileIds || {})[`${vLower}||${fileName}`];
    return approvedId || null;
  }


  public _renderBentoComposeModal(): React.ReactElement | React.ReactPortal | null {
    // "Send Email" now opens the Outlook-style composer (To/Cc/Bcc, any
    // SharePoint files/folders, files from the computer). The older AI
    // Bento dispatch form (renderBentoComposeModal) is no longer opened.
    if (!this.state.bentoComposeOpen) return null;
    return <ComposeMailModal host={this} onClose={() => this.setState({ bentoComposeOpen: false })} />;
  }


  // ── Existing Views: Recycle Bin ───────────────────────────────────────────

  /**
   * Helper to resolve the most accurate SharePoint site display name for any DeletedNode,
   * checking explicit properties, site_key match against documentSites, vessel associations, or active site fallback.
   */
  public _resolveSiteNameForDeletedNode = (node: DeletedNode): string => {
    if (node.site_name && node.site_name.trim()) return node.site_name.trim();

    // 1. Try matching node.site_key
    if (node.site_key) {
      const matched = (this.state.documentSites || []).find(s =>
        s.site_key === node.site_key || s.site_id === node.site_key || s.sp_site_name === node.site_key
      );
      if (matched?.sp_site_name || matched?.site_key) {
        return matched.sp_site_name || matched.site_key;
      }
    }

    // 2. If node is a vessel, check provisioned sites of matching vessel
    const isVesselKind = node.kind === 'vessel' || node.item_type === 'vessel';
    const targetVesselName = (isVesselKind ? node.name : node.vessel_name || '').trim().toLowerCase();
    if (targetVesselName) {
      const v = (this.state.vessels || []).find(ves => cleanName(ves.name).trim().toLowerCase() === cleanName(targetVesselName).trim().toLowerCase());
      if (v && v.provisioned_site_ids && v.provisioned_site_ids.length > 0) {
        const names = v.provisioned_site_ids.map(sk => {
          const m = (this.state.documentSites || []).find(s => s.site_key === sk || s.site_id === sk);
          return m?.sp_site_name || sk;
        }).filter(Boolean);
        if (names.length > 0) return names.join(', ');
      }
    }

    // 3. Check if path indicates a specific site
    const pathLower = (node.original_path || '').toLowerCase();
    for (const site of (this.state.documentSites || [])) {
      if (site.sp_site_name && pathLower.includes(site.sp_site_name.toLowerCase())) {
        return site.sp_site_name;
      }
      if (site.site_key && pathLower.includes(site.site_key.toLowerCase())) {
        return site.sp_site_name || site.site_key;
      }
    }

    // 4. Default to current active site or siteTitle
    const activeSiteObj = (this.state.documentSites || []).find(s => s.site_key === this.state.activeDocumentSite);
    return activeSiteObj?.sp_site_name || this.props.siteUrl || 'SharePoint';
  };

  /**
   * Load all deleted items for the Recycle Bin page:
   * 1. Fetches backend deleted nodes database records.
   * 2. Queries SharePoint Online Recycle Bin (web and site collection) across all configured sites.
   * 3. Merges and deduplicates with active in-memory session deletions, populating site_name and site_key.
   */
  public _loadRecycleBin = async (): Promise<void> => {
    this.setState({ panelLoading: true });
    try {
      let backendItems: DeletedNode[] = [];
      try {
        const data = await this._fetchJson(`${this._base()}/api/recycle-bin/nodes`);
        backendItems = (data || []).map((v: any) => ({ ...v, name: cleanName(v.name || '') }));
      } catch {}

      const spoRecycledNodes: DeletedNode[] = [];

      // Collect all configured SharePoint sites to query (primary site from props + documentSites)
      const sitesToQuery: Array<{ url: string; siteKey?: string; siteName?: string }> = [];
      const primarySiteUrl = this.props.siteUrl;
      if (primarySiteUrl) {
        const primarySiteObj = (this.state.documentSites || []).find(s =>
          s.is_primary || (s.web_url && s.web_url.toLowerCase() === primarySiteUrl.toLowerCase())
        );
        sitesToQuery.push({
          url: primarySiteUrl.replace(/\/+$/, ''),
          siteKey: primarySiteObj?.site_key,
          siteName: primarySiteObj?.sp_site_name || this.props.siteUrl || 'SharePoint',
        });
      }
      for (const ds of (this.state.documentSites || [])) {
        const wUrl = (ds.web_url || '').replace(/\/+$/, '');
        if (wUrl && !sitesToQuery.some(s => s.url.toLowerCase() === wUrl.toLowerCase())) {
          sitesToQuery.push({
            url: wUrl,
            siteKey: ds.site_key,
            siteName: ds.sp_site_name || ds.site_key,
          });
        }
      }

      for (const targetSite of sitesToQuery) {
        const endpoints = [
          `${targetSite.url}/_api/web/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType,DeletedDate,DeletedByName&$top=500`,
          `${targetSite.url}/_api/site/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType,DeletedDate,DeletedByName&$top=500`,
        ];
        for (const ep of endpoints) {
          try {
            const res = await fetch(ep, {
              headers: {
                Accept: 'application/json;odata=nometadata, application/json;odata=verbose, application/json',
              },
            });
            if (res.ok) {
              const resData = await res.json();
              const items = resData?.value || resData?.d?.results || (Array.isArray(resData) ? resData : []);
              for (const it of items) {
                const id = it.Id || it.id;
                if (!id) continue;
                if (spoRecycledNodes.some(n => n.recycle_bin_item_id === id || n.id === id)) continue;

                const leafName = it.LeafName || it.leafName || it.Title || it.title || '';
                if (!leafName) continue;

                const cleanDir = (it.DirName || it.dirName || '')
                  .replace(/^Shared Documents\/?/i, '')
                  .replace(/^\/+|\/+$/g, '');
                const dirParts = cleanDir.split('/').filter(Boolean);
                const mainFolder = dirParts.find((p: string) => this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === p.toLowerCase())) || 'Technical & Crewing';
                const vesselName = dirParts.find((p: string) => !this.MAIN_FOLDER_NAMES.some(mf => mf.toLowerCase() === p.toLowerCase()) && !/^(vessels|specific vessels|documents)$/i.test(p)) || '';
                const originalPath = cleanDir ? `${cleanDir}/${leafName}` : leafName;

                const isFileItem = it.ItemType === 1 || it.ItemType === '1' || it.ItemType === 'File' || it.ItemType === 'FileItem';
                const isFolderItem = it.ItemType === 2 || it.ItemType === 5 || it.ItemType === '2' || it.ItemType === 'Folder' || it.ItemType === 'FolderItem';
                const hasFileExt = /\.[a-zA-Z0-9]{1,8}$/.test(leafName);

                let isFolder = false;
                if (isFolderItem) {
                  isFolder = true;
                } else if (isFileItem) {
                  isFolder = false;
                } else {
                  isFolder = !hasFileExt;
                }

                const isVesselFolder = (
                  cleanDir.toLowerCase() === 'vessels/specific vessels' ||
                  cleanDir.toLowerCase() === 'vessels' ||
                  (cleanDir.toLowerCase().endsWith('specific vessels') && !cleanDir.toLowerCase().includes('drawing') && !cleanDir.toLowerCase().includes('manual'))
                );

                const kind = isVesselFolder ? 'vessel' : (isFolder ? 'folder' : 'file');
                const itemType = isVesselFolder ? 'vessel' : (isFolder ? 'folder' : 'file');

                spoRecycledNodes.push({
                  id: String(id),
                  name: cleanName(leafName),
                  kind,
                  item_type: itemType,
                  main_folder: mainFolder,
                  vessel_name: isVesselFolder ? cleanName(leafName) : vesselName,
                  original_path: originalPath,
                  deleted_at: it.DeletedDate || it.deletedDate || new Date().toISOString(),
                  in_spo_recycle_bin: true,
                  recycle_bin_item_id: String(id),
                  site_name: targetSite.siteName,
                  site_key: targetSite.siteKey,
                });
              }
            }
          } catch (err) {
            console.warn(`[VesselDMS] Fetching SPO RecycleBin items warning for ${targetSite.url}:`, err);
          }
        }
      }

      this.setState(prev => {
        const prevLocal = prev.recycleBin || [];
        const combined = [...backendItems];
        const sameRecycleItem = (left: DeletedNode, right: DeletedNode): boolean => {
          const sameSite = !left.site_key || !right.site_key || left.site_key === right.site_key;
          return sameSite && (
            left.id === right.id ||
            (!!left.recycle_bin_item_id && left.recycle_bin_item_id === right.recycle_bin_item_id) ||
            (left.name.toLowerCase() === right.name.toLowerCase() && left.original_path === right.original_path)
          );
        };

        // Add SPO recycled nodes not already present in combined
        for (const spoNode of spoRecycledNodes) {
          const exists = combined.some(b => sameRecycleItem(b, spoNode));
          if (!exists) {
            combined.push(spoNode);
          } else {
            const match = combined.find(b => sameRecycleItem(b, spoNode));
            if (match && spoNode.site_name) {
              match.site_name = spoNode.site_name;
              match.site_key = spoNode.site_key || match.site_key;
            }
          }
        }

        // Add locally tracked items not in combined
        for (const locNode of prevLocal) {
          const exists = combined.some(c => sameRecycleItem(c, locNode));
          if (!exists) {
            combined.push(locNode);
          } else {
            const match = combined.find(c => sameRecycleItem(c, locNode));
            if (match && locNode.site_name) {
              match.site_name = locNode.site_name;
              match.site_key = locNode.site_key || match.site_key;
            }
          }
        }

        // Backfill site_name for any nodes that still lack it
        for (const node of combined) {
          if (!node.site_name) {
            node.site_name = this._resolveSiteNameForDeletedNode(node);
          }
        }

        return { recycleBin: combined, panelLoading: false };
      });
    } catch {
      this.setState({ panelLoading: false });
    }
  };

  /**
   * Restore a vessel folder, normal folder, or file from the SPO site Recycle Bin
   * back to its original SharePoint location using the SharePoint REST & Graph APIs.
   */
  public async _restoreSpoItem(item: DeletedNode): Promise<{ success: boolean; restoredCount: number; message?: string }> {
    if (!this.props.siteUrl) return { success: false, restoredCount: 0, message: 'No site URL configured' };
    const cleanItemName = cleanName(item.name || '').replace(/\/+$/, '').trim().toLowerCase();
    const rawItemName = (item.name || '').replace(/\/+$/, '').trim().toLowerCase();
    const isVessel = item.kind === 'vessel' || item.item_type === 'vessel';

    // Resolve target site URL and site ID
    let targetSiteUrl = this.props.siteUrl;
    let targetSiteId = this.props.siteId;
    if (item.site_key || item.site_name) {
      const matchedSite = (this.state.documentSites || []).find(s =>
        (item.site_key && (s.site_key === item.site_key || s.site_id === item.site_key)) ||
        (item.site_name && (s.sp_site_name?.toLowerCase() === item.site_name.toLowerCase() || s.site_key?.toLowerCase() === item.site_name.toLowerCase()))
      );
      if (matchedSite?.web_url) targetSiteUrl = matchedSite.web_url;
      if (matchedSite?.site_id) targetSiteId = matchedSite.site_id;
    }

    try {
      // 1. Get CSRF digest
      let digest = '';
      try {
        const digestRes = await fetch(`${targetSiteUrl}/_api/contextinfo`, {
          method: 'POST',
          headers: { Accept: 'application/json;odata=nometadata' },
        });
        const digestData = await digestRes.json();
        digest = digestData?.FormDigestValue ?? '';
      } catch (dErr) {
        console.warn('[VesselDMS] Could not fetch request digest:', dErr);
      }

      // 2. Query both web and site Recycle Bin collections without $filter (which is unsupported on SPRecycleBin)
      const allRecycledItems: any[] = [];
      const endpointsToQuery = [
        `${targetSiteUrl}/_api/web/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType&$top=500`,
        `${targetSiteUrl}/_api/site/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType&$top=500`,
      ];

      for (const endpoint of endpointsToQuery) {
        try {
          const rbRes = await fetch(endpoint, {
            headers: { Accept: 'application/json;odata=nometadata' },
          });
          if (rbRes.ok) {
            const rbData = await rbRes.json();
            const items = rbData?.value ?? [];
            for (const it of items) {
              if (it.Id && !allRecycledItems.some(existing => existing.Id === it.Id)) {
                allRecycledItems.push(it);
              }
            }
          }
        } catch (fetchErr) {
          console.warn(`[VesselDMS] Querying ${endpoint} warning:`, fetchErr);
        }
      }

      // 3. Match items by LeafName, Title, or sanitized/cleaned names
      const matchNorm = (s: string) => cleanName(s || '').replace(/\/+$/, '').trim().toLowerCase();
      const noSpace = (s: string) => (s || '').replace(/[\s_\-'\"]+/g, '').toLowerCase();

      const matchingItems = allRecycledItems.filter(r => {
        const leaf = matchNorm(r.LeafName);
        const title = matchNorm(r.Title);
        return (
          leaf === cleanItemName ||
          leaf === rawItemName ||
          title === cleanItemName ||
          title === rawItemName ||
          noSpace(r.LeafName) === noSpace(item.name) ||
          (isVessel && (leaf.includes(cleanItemName) || cleanItemName.includes(leaf)))
        );
      });

      console.log(`[VesselDMS] _restoreSpoItem found ${matchingItems.length} matching item(s) in SharePoint Recycle Bin for "${item.name}"`);

      // 4. Restore each matching item back to its original location
      let restoredCount = 0;
      for (const match of matchingItems) {
        const guid = match.Id;
        if (!guid) continue;

        let restoredThis = false;
        // Try web and site restore endpoints with both guid'...' and string syntax
        const restoreCalls = [
          `${targetSiteUrl}/_api/web/RecycleBin(guid'${guid}')/restore()`,
          `${targetSiteUrl}/_api/site/RecycleBin(guid'${guid}')/restore()`,
          `${targetSiteUrl}/_api/web/RecycleBin('${guid}')/restore()`,
          `${targetSiteUrl}/_api/site/RecycleBin('${guid}')/restore()`,
        ];

        for (const restoreUrl of restoreCalls) {
          try {
            const res = await fetch(restoreUrl, {
              method: 'POST',
              headers: {
                Accept: 'application/json;odata=nometadata',
                'X-RequestDigest': digest,
              },
            });
            if (res.ok || res.status === 204) {
              restoredThis = true;
              console.log(`[VesselDMS] Successfully restored SPO item ${guid} (${match.DirName}/${match.LeafName}) via ${restoreUrl}`);
              break;
            }
          } catch (rErr) {
            // Try next syntax
          }
        }

        // Try Graph API as fallback if available
        if (!restoredThis && this.props.graphClient && targetSiteId) {
          try {
            await this.props.graphClient
              .api(`/sites/${targetSiteId}/recycleBin/items/${guid}/restore`)
              .post({});
            restoredThis = true;
            console.log(`[VesselDMS] Successfully restored SPO item ${guid} via Graph API`);
          } catch (gErr) {
            // Ignore fallback error
          }
        }

        if (restoredThis) {
          restoredCount++;
        }
      }

      // 5. (Removed) This used to call _provisionVesselFolders(item.name) as a
      // "just in case" fallback to recreate the vessel's folder structure.
      // With no vesselId/targetSiteIds available here, _provisionVesselFolders
      // always fell through to the legacy client-side createVesselFolders()
      // path (graphFolderService.ts), which builds the OLD multi-department
      // nested tree (Technical & Crewing / Commercial & Chartering / Insurance
      // + full perVesselTree) — not the single flat root folder this app now
      // uses. Because createFolder() treats a 409 as "already exists" and
      // just resolves it, this silently created the old department folders
      // in real SharePoint on every recycle-bin restore. Step 4 above already
      // restores whatever was actually deleted from the Recycle Bin, which is
      // the correct/complete behavior — no additional provisioning fallback
      // is needed or wanted.

      return {
        success: true,
        restoredCount,
        message: restoredCount > 0
          ? `Restored ${restoredCount} item(s) from SharePoint Recycle Bin.`
          : 'Reactivated item and restored folder tree.',
      };
    } catch (err: any) {
      console.warn('[VesselDMS] _restoreSpoItem failed:', err);
      return { success: false, restoredCount: 0, message: err?.message ?? String(err) };
    }
  }

  public _restoreFromRecycleBin = async (item: DeletedNode): Promise<{ ok: boolean; message?: string }> => {
    const isVessel = item.kind === 'vessel' || item.item_type === 'vessel';

    // 1. Restore item(s) from SharePoint Online site Recycle Bin back to their original paths
    try {
      await this._restoreSpoItem(item);
    } catch (spoErr) {
      console.warn('[VesselDMS] SPO restore step warning:', spoErr);
    }

    // 2. Notify backend to restore database record
    try {
      const query = new URLSearchParams({
        type: item.kind === 'file' ? 'file' : 'folder',
        item_name: item.name,
        department: item.main_folder || '',
        vessel_name: item.vessel_name || '',
        site_key: item.site_key || '',
        site_name: item.site_name || '',
      });
      const res = await fetch(`${this._base()}/api/recycle-bin/restore/${encodeURIComponent(item.id)}?${query.toString()}`, {
        method: 'POST', headers: this._headers(),
      });
      if (!res.ok && res.status !== 202) {
        console.warn(`Backend restore returned status ${res.status}`);
      }
    } catch (e: any) {
      console.warn('[VesselDMS] Backend restore error:', e);
    }

    // 3. If it's a vessel, restore vessel record into active vessels state
    if (isVessel) {
      const base = this._base();
      let restoredId = item.id.replace(/^vessel_/, '').replace(/^db_vessel_/, '');
      try {
        const created = await this._fetchJson(`${base}/api/vessels`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: item.name,
            imo: item.imo && item.imo !== '—' ? item.imo : (item.vessel_imo && item.vessel_imo !== '—' ? item.vessel_imo : null),
            shipyard: item.shipyard || 'Restored',
            vessel_type: item.vessel_type || 'Bulk Carrier',
            site_key: item.site_key || undefined,
            site_name: item.site_name || undefined,
          }),
        });
        if (created?.id) restoredId = String(created.id);
      } catch { /* already exists or offline fallback */ }

      const restoredVessel: VesselRecord = {
        id: restoredId,
        name: cleanName(item.name),
        imo: item.imo || item.vessel_imo || '—',
        shipyard: item.shipyard || 'Restored',
        hull_number: item.hull_number || '',
        vessel_type: item.vessel_type || 'Bulk Carrier',
        status: 'Active',
        image_url: pickRandomVesselImage(item.vessel_type || 'Bulk Carrier'),
        is_provisioned: true,
        provisioned_site_ids: item.site_key ? [item.site_key] : [],
      };

      // Remove from spoDeletedVesselIds so it's no longer filtered out of vessel lists
      this.setState(prev => ({
        vessels: [restoredVessel, ...prev.vessels.filter(v => v.name.toLowerCase() !== restoredVessel.name.toLowerCase())],
        spoDeletedVesselIds: (() => { const s = new Set(prev.spoDeletedVesselIds); s.delete(item.id); s.delete(restoredId); return s; })(),
      }));
    }

    // Remove from local Recycle Bin state
    this.setState(prev => ({ recycleBin: prev.recycleBin.filter(r => r.id !== item.id) }));

    // Trigger delta sync so the restored SPO folder/file appears in the tree
    void this._syncScheduler?.triggerNow().catch(() => undefined);

    const destPath = item.original_path || (isVessel ? `All Departments / ${item.name}` : item.name);
    return { ok: true, message: `Restored to ${destPath}` };
  };

  public _permanentDeleteFromRecycleBin = async (item: DeletedNode, skipConfirm?: boolean): Promise<{ ok: boolean; message?: string }> => {
    if (!skipConfirm && !window.confirm(`Permanently delete "${item.name}"? This action cannot be undone.`)) return { ok: false, message: 'Deletion cancelled.' };

    const { graphClient, siteId, driveId } = this.props;
    const isRealGraphId = item.id && !/^\d+$/.test(item.id) && !/^file_/.test(item.id) &&
      !/^vessel_/.test(item.id) && !/^anomaly_/.test(item.id) && !/^db_vessel_/.test(item.id);

    // Resolve target site URL
    let targetSiteUrl = this.props.siteUrl;
    if (item.site_key || item.site_name) {
      const matchedSite = (this.state.documentSites || []).find(s =>
        (item.site_key && (s.site_key === item.site_key || s.site_id === item.site_key)) ||
        (item.site_name && (s.sp_site_name?.toLowerCase() === item.site_name.toLowerCase() || s.site_key?.toLowerCase() === item.site_name.toLowerCase()))
      );
      if (matchedSite?.web_url) targetSiteUrl = matchedSite.web_url;
    }

    try {
      if (graphClient && siteId && driveId && isRealGraphId) {
        try {
          // Graph /recycleBin is not supported in v1.0 - use SharePoint REST API instead.
          const findInRecycleBin = async (): Promise<string | null> => {
            if (!targetSiteUrl) return null;
            try {
              const nameLower = item.name.toLowerCase();
              const rbRes = await fetch(
                `${targetSiteUrl}/_api/site/RecycleBin?$filter=LeafName eq '${encodeURIComponent(item.name)}'&$select=Id,LeafName&$top=10`,
                { headers: { Accept: 'application/json;odata=nometadata' } }
              );
              if (!rbRes.ok) return null;
              const rbData = await rbRes.json();
              const match = (rbData?.value ?? []).find((r: any) => (r.LeafName || '').toLowerCase() === nameLower);
              return match?.Id ?? null;
            } catch { return null; }
          };

          // Use SharePoint REST API to delete from recycle bin (Graph /recycleBin not supported in v1.0).
          // SPO REST mutating calls require X-RequestDigest for CSRF validation.
          const getSpoDigest = async (): Promise<string> => {
            const r = await fetch(`${targetSiteUrl}/_api/contextinfo`, {
              method: 'POST',
              headers: { Accept: 'application/json;odata=nometadata' },
            });
            const d = await r.json();
            return d?.FormDigestValue ?? d?.['odata.metadata'] ?? '';
          };
          const deleteFromRecycleBin = async (rbId: string): Promise<void> => {
            if (!targetSiteUrl) return;
            const digest = await getSpoDigest();
            const res = await fetch(
              `${targetSiteUrl}/_api/site/RecycleBin('${rbId}')`,
              { method: 'POST', headers: { Accept: 'application/json;odata=nometadata', 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*', 'X-RequestDigest': digest } }
            );
            if (!res.ok && res.status !== 204) throw new Error(`RecycleBin delete failed (${res.status})`);
          };

          if (item.in_spo_recycle_bin && item.recycle_bin_item_id) {
            // Fast path: we already have the recycle bin item ID.
            await deleteFromRecycleBin(item.recycle_bin_item_id);
          } else {
            // Always check the SPO recycle bin first (item may be there regardless of flag).
            const rbItemId = await findInRecycleBin();
            if (rbItemId) {
              await deleteFromRecycleBin(rbItemId);
            } else {
              // Not found in SPO recycle bin — call backend to hard-delete the DB record.
              console.warn('[VesselDMS] permanentDelete: item not found in SPO recycle bin, falling back to backend delete:', item.name);
              const query = new URLSearchParams({
                type: item.kind === 'file' ? 'file' : 'folder',
                item_name: item.name,
                department: item.main_folder || '',
                vessel_name: item.vessel_name || '',
                site_key: item.site_key || '',
                site_name: item.site_name || '',
              });
              const res = await fetch(`${this._base()}/api/recycle-bin/${encodeURIComponent(item.id)}?${query.toString()}`, {
                method: 'DELETE', headers: this._headers(),
              });
              const data = await res.json().catch(() => ({}));
              if (!res.ok && res.status !== 202) throw new Error(data?.message || `Permanent delete failed (${res.status})`);
              if (res.status === 202) {
                return { ok: false, message: data?.message || 'Deletion was submitted for approval.' };
              }
            }
          }
        } catch (graphErr: any) {
          // 404 / itemNotFound = already gone, treat as success.
          if (graphErr?.statusCode !== 404 && graphErr?.code !== 'itemNotFound') {
            throw graphErr;
          }
        }
      } else {
        // Fallback: backend hard-delete (vessels, folders, or items without a Graph ID)
        const query = new URLSearchParams({
          type: item.kind === 'file' ? 'file' : 'folder',
          item_name: item.name,
          department: item.main_folder || '',
          vessel_name: item.vessel_name || '',
          site_key: item.site_key || '',
          site_name: item.site_name || '',
        });
        const res = await fetch(`${this._base()}/api/recycle-bin/${encodeURIComponent(item.id)}?${query.toString()}`, {
          method: 'DELETE', headers: this._headers(),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok && res.status !== 202) throw new Error(data?.message || `Permanent delete failed (${res.status})`);
        if (res.status === 202) {
          return { ok: false, message: data?.message || 'Deletion was submitted for approval.' };
        }
      }
    } catch (e: any) {
      return { ok: false, message: e?.message || 'Could not permanently delete this item.' };
    }
    this.setState(prev => ({ recycleBin: prev.recycleBin.filter(r => r.id !== item.id) }));
    void this._syncScheduler?.triggerNow().catch(() => undefined);
    return { ok: true };
  };


  // ── SharePoint Recycle Bin Folder ────────────────────────────────────────

  /**
   * Ensure Documents/Recycle Bin/ exists in SharePoint, then move the given
   * drive item into it.  Called after every successful delete so the physical
   * SPO folder mirrors the DB recycle-bin state.
   *
   * Graph move = PATCH /drives/{driveId}/items/{itemId}
   *   { parentReference: { id: <recycleBinFolderId> }, name: <name> }
   */
  /**
   * Soft-delete a drive item via Graph DELETE — moves it to the SPO site Recycle Bin.
   */
  public async _moveToSharePointRecycleBin(
    itemId: string,
    _itemName: string,
  ): Promise<void> {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId || !itemId) return;
    try {
      await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${itemId}`).delete();
      void this._syncScheduler?.triggerNow().catch(() => undefined);
    } catch (err) {
      console.warn('[VesselDMS] _moveToSharePointRecycleBin (soft-delete) failed:', err);
    }
  }

  public _renderRecycleBinPage(): React.ReactElement {
    return renderRecycleBinPage(this);
  }

  // Self-contained module — see components/migrationAssistant/README.md.
  // Unlike the render* methods above (which take `host` and read/write the
  // root component's own state), this module owns its own state entirely;
  // the root component only supplies which backend to talk to and as whom.
  public _renderMigrationAssistant(): React.ReactElement {
    return (
      <PageErrorBoundary pageName="Migration Assistant">
        <MigrationAssistantModule
          apiBaseUrl={this.props.migrationApiBaseUrl || this.props.apiBaseUrl || ''}
          sessionId={this.props.sessionId}
          actingEmail={this.props.userEmail}
          isNight={this.state.themeMode === 'night'}
        />
      </PageErrorBoundary>
    );
  }

  public _renderSpoVesselDeletedToast(): React.ReactElement | null {
    const { spoVesselDeletedToast } = this.state;
    if (!spoVesselDeletedToast) return null;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
    const { vesselName, siteName, originalPath, deletedByName, deletedByEmail } = spoVesselDeletedToast;
    const deletedBy = deletedByName || deletedByEmail || 'Unknown user (pending SharePoint confirmation)';
    return (
      <div
        onMouseEnter={this._pauseVesselToastAutoClose}
        onMouseLeave={this._scheduleVesselToastAutoClose}
        style={{
          pointerEvents: 'auto',
          background: '#1e293b', color: '#fff',
          borderRadius: 12, boxShadow: '0 8px 32px rgba(0,0,0,0.28)',
          padding: isMobile ? '12px' : '16px 24px', display: 'flex', alignItems: 'center', gap: 12,
          flexDirection: isMobile ? 'column' : 'row', width: isMobile ? 'calc(100vw - 20px)' : 'auto', minWidth: isMobile ? 0 : 360, maxWidth: isMobile ? 'calc(100vw - 20px)' : 560, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
          border: '1.5px solid #ef4444',
        }}>
        <div style={{
          width: 36, height: 36, borderRadius: 10, background: '#fee2e2',
          display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0,
        }}><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 18 }} /></div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: '#fca5a5' }}>Vessel deleted in SharePoint</div>
          <div style={{ fontSize: 12, color: 'var(--vdms-text-faint)', marginTop: 3 }}>
            <strong style={{ color: '#fff' }}>{vesselName}</strong> was removed from SharePoint Online and has been moved to the Recycle Bin.
          </div>
          <div style={{ marginTop: 8, display: 'grid', gap: 3, fontSize: 11, color: 'var(--vdms-text-faint)' }}>
            <div><strong style={{ color: '#fff' }}>Deleted by:</strong> {deletedBy}</div>
            {siteName && <div><strong style={{ color: '#fff' }}>Site:</strong> {siteName}</div>}
            {originalPath && <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}><strong style={{ color: '#fff' }}>Path:</strong> {originalPath}</div>}
          </div>
        </div>
        <div style={{ display: 'flex', flexDirection: isMobile ? 'row' : 'column', gap: 6, flexShrink: 0, width: isMobile ? '100%' : 'auto' }}>
          <button
            onClick={() => {
              this._pauseVesselToastAutoClose();
              this.setState({ spoVesselDeletedToast: null });
              void this._goToView('recycle');
            }}
            style={{
              background: '#ef4444', color: '#fff', border: 'none', borderRadius: 7,
              minHeight: 44, flex: isMobile ? 1 : undefined, padding: '5px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
            }}
          >View Recycle Bin</button>
          <button
            onClick={() => { this._pauseVesselToastAutoClose(); this.setState({ spoVesselDeletedToast: null }); }}
            style={{
              background: 'transparent', color: 'var(--vdms-text-faint)', border: '1px solid var(--vdms-text-secondary)',
              minHeight: 44, flex: isMobile ? 1 : undefined, borderRadius: 7, padding: '5px 12px', fontSize: 11, fontWeight: 600, cursor: 'pointer',
            }}
          >Dismiss</button>
        </div>
      </div>
    );
  }

  public _renderSpoDocumentDeletedToast(): React.ReactElement | null {
    const { spoDocumentDeletedToast } = this.state;
    if (!spoDocumentDeletedToast) return null;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
    const { items, itemType } = spoDocumentDeletedToast;
    const label = itemType === 'folder' ? 'Folder' : 'File';
    const summary = items.length === 1
      ? items[0].name
      : `${items.length} ${itemType}s`;
    // If every item in this batch shares the same actor/site, show it once
    // as a summary line; otherwise fall back to a per-item breakdown so
    // nothing gets misattributed to "the" deleter.
    const uniqueBy = <T,>(pick: (i: typeof items[number]) => T | null | undefined): (T | null | undefined)[] =>
      Array.from(new Set(items.map(pick)));
    const deleters = uniqueBy(i => i.deletedByName || i.deletedByEmail || null);
    const sites = uniqueBy(i => i.siteName || null);
    const singleDeleter = deleters.length === 1 ? deleters[0] : null;
    const singleSite = sites.length === 1 ? sites[0] : null;
    return (
      <div
        onMouseEnter={this._pauseDocumentToastAutoClose}
        onMouseLeave={this._scheduleDocumentToastAutoClose}
        style={{
          pointerEvents: 'auto',
          background: '#1e293b', color: '#fff', borderRadius: 12,
          boxShadow: '0 8px 32px rgba(0,0,0,0.28)', padding: isMobile ? '12px' : '16px 24px',
          display: 'flex', alignItems: 'flex-start', gap: 12, flexDirection: isMobile ? 'column' : 'row', width: isMobile ? 'calc(100vw - 20px)' : 'auto', minWidth: isMobile ? 0 : 360, maxWidth: isMobile ? 'calc(100vw - 20px)' : 560,
          fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", border: '1.5px solid #ef4444',
        }}>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0 }}><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 18 }} /></div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: '#fca5a5' }}>{items.length === 1 ? label : `${items.length} ${itemType}s`} moved to Recycle Bin</div>
          <div style={{ fontSize: 12, color: 'var(--vdms-text-faint)', marginTop: 3 }}>
            <strong style={{ color: '#fff' }}>{summary}</strong> {items.length === 1 ? 'was' : 'were'} deleted from SharePoint Online and moved to the Recycle Bin.
          </div>

          {items.length === 1 ? (
            <div style={{ marginTop: 8, display: 'grid', gap: 3, fontSize: 11, color: 'var(--vdms-text-faint)' }}>
              <div><strong style={{ color: '#fff' }}>Deleted by:</strong> {items[0].deletedByName || items[0].deletedByEmail || 'Unknown user (pending SharePoint confirmation)'}</div>
              {items[0].siteName && <div><strong style={{ color: '#fff' }}>Site:</strong> {items[0].siteName}</div>}
              {items[0].path && <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}><strong style={{ color: '#fff' }}>Path:</strong> {items[0].path}</div>}
            </div>
          ) : (
            <>
              {(singleDeleter || singleSite) && (
                <div style={{ marginTop: 8, display: 'grid', gap: 3, fontSize: 11, color: 'var(--vdms-text-faint)' }}>
                  {singleDeleter && <div><strong style={{ color: '#fff' }}>Deleted by:</strong> {singleDeleter}</div>}
                  {singleSite && <div><strong style={{ color: '#fff' }}>Site:</strong> {singleSite}</div>}
                </div>
              )}
              <div style={{ marginTop: 8, maxHeight: 140, overflowY: 'auto', border: '1px solid rgba(255,255,255,0.12)', borderRadius: 8 }}>
                {items.map(entry => (
                  <div key={entry.id} style={{ padding: '6px 8px', borderBottom: '1px solid rgba(255,255,255,0.08)', fontSize: 11, color: 'var(--vdms-text-faint)' }}>
                    <div style={{ color: '#fff', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{entry.name}</div>
                    <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {!singleDeleter && `By ${entry.deletedByName || entry.deletedByEmail || 'Unknown user'} · `}
                      {!singleSite && entry.siteName ? `${entry.siteName} · ` : ''}
                      {entry.path || '—'}
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
        <div style={{ display: 'flex', flexDirection: isMobile ? 'row' : 'column', gap: 6, flexShrink: 0, width: isMobile ? '100%' : 'auto' }}>
          <button onClick={() => { this._pauseDocumentToastAutoClose(); this.setState({ spoDocumentDeletedToast: null }); void this._goToView('recycle'); }} style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 7, minHeight: 44, flex: isMobile ? 1 : undefined, padding: '5px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}>View Recycle Bin</button>
          <button onClick={() => { this._pauseDocumentToastAutoClose(); this.setState({ spoDocumentDeletedToast: null }); }} style={{ background: 'transparent', color: 'var(--vdms-text-faint)', border: '1px solid var(--vdms-text-secondary)', borderRadius: 7, minHeight: 44, flex: isMobile ? 1 : undefined, padding: '5px 12px', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>Dismiss</button>
        </div>
      </div>
    );
  }

  /** Shared fixed-position stack for the two SPO-deletion popups so a
   *  vessel-deletion toast and a document-deletion toast that happen to be
   *  active at the same time stack cleanly above one another instead of
   *  both being pinned to the exact same spot and overlapping. */
  public _renderSpoDeletionToasts(): React.ReactElement | null {
    const { spoVesselDeletedToast, spoDocumentDeletedToast } = this.state;
    if (!spoVesselDeletedToast && !spoDocumentDeletedToast) return null;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
    return (
      <div style={{
        position: 'fixed', bottom: isMobile ? 10 : 28, left: '50%', transform: 'translateX(-50%)',
        zIndex: 100003, display: 'flex', flexDirection: 'column', gap: 10,
        alignItems: isMobile ? 'stretch' : 'center', pointerEvents: 'none',
      }}>
        {this._renderSpoVesselDeletedToast()}
        {this._renderSpoDocumentDeletedToast()}
      </div>
    );
  }

  public _renderVesselForm(mode: 'create' | 'edit'): React.ReactElement {
    return renderVesselForm(this, mode);
  }

  public _renderDeleteModal(): React.ReactElement {
    return renderDeleteModal(this);
  }

  public _openBulkUpload = (
    files: BulkUploadFile[],
    folderId: string,
    subFolderPath: string,
    vesselName: string,
    currentFolderNode?: { id: string; name: string } | null,
    targetSiteId?: string,
    targetDriveId?: string,
  ): void => {
    console.info('[VesselDMS] _openBulkUpload', { count: files ? files.length : 0, folderId, subFolderPath, vesselName });
    if (!files || files.length === 0) return;
    this.setState({
      bulkUploadDialog: {
        files,
        folderId,
        subFolderPath,
        vesselName,
        currentFolderNode: currentFolderNode || null,
        targetSiteId,
        targetDriveId,
      },
    });
  };

  public _renderBulkUploadModal(): React.ReactElement | null {
    const { bulkUploadDialog } = this.state;
    if (!bulkUploadDialog) return null;
    return (
      <BulkUploadModal
        host={this}
        files={bulkUploadDialog.files}
        folderId={bulkUploadDialog.folderId}
        subFolderPath={bulkUploadDialog.subFolderPath}
        vesselName={bulkUploadDialog.vesselName}
        currentFolderNode={bulkUploadDialog.currentFolderNode}
        targetSiteId={bulkUploadDialog.targetSiteId}
        targetDriveId={bulkUploadDialog.targetDriveId}
        onClose={() => { console.info('[VesselDMS] bulk upload dialog onClose', new Error().stack); this.setState({ bulkUploadDialog: null }); }}
      />
    );
  }

  public render(): React.ReactElement {
    const { sessionExpired, view, authPage } = this.state;

    // Keep <body> in sync with the active theme. The --vdms-* custom
    // properties consumed by inline styles throughout the app (including
    // this file's dialog renderers such as _renderFileDeleteDialog,
    // _renderAddFolderDialog, and _renderFolderDeleteDialog) are only
    // defined under the [data-vessel-theme="light"|"night"] selector in
    // injectFullScreenStyles. That attribute is set on the themed
    // .vessel-dms-app wrapper (see AppLayout.tsx renderLayout), but these
    // dialogs are rendered as top-level siblings of that wrapper — not as
    // its descendants — so the vars previously resolved to nothing there,
    // leaving the popups with a transparent/undefined background and text
    // color that blended into the backdrop overlay. Mirroring the
    // attribute onto <body> makes the tokens resolve everywhere.
    if (typeof document !== 'undefined') {
      document.body.setAttribute('data-vessel-theme', this.state.themeMode);
    }

    if (authPage) {
      return renderAuthPage(this, authPage);
    }

    if (sessionExpired) {
      const isNight = this.state.themeMode === 'night';
      return (
        <div className="vessel-dms-auth" data-vessel-theme={this.state.themeMode} style={{
          position: 'fixed', inset: 0, zIndex: 99999, isolation: 'isolate',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          background: clay.bg, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", padding: 24,
        }}>
          <div style={{
            background: isNight ? 'rgba(7,28,47,0.9)' : 'rgba(255,255,255,0.92)', backdropFilter: 'blur(26px) saturate(1.5)', WebkitBackdropFilter: 'blur(26px) saturate(1.5)',
            borderRadius: 28, padding: '44px 48px', maxWidth: 460,
            textAlign: 'center', boxShadow: 'var(--vdms-shadow)', border: '1px solid var(--vdms-line)',
          }}>
            <div style={{ width: 64, height: 64, margin: '0 auto 18px', borderRadius: 20, background: clay.accentGradient, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 30, boxShadow: clay.shadowButton }}><Icon iconName="Lock" aria-hidden="true" style={{ fontSize: 30 }} /></div>
            <h2 style={{ margin: '0 0 10px', fontSize: 24, fontWeight: 800, color: 'var(--vdms-text)', fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}>Session Expired</h2>
            <p style={{ margin: '0 0 26px', fontSize: 15, color: 'var(--vdms-text-secondary)', lineHeight: 1.6 }}>
              Your session has expired or is no longer valid. Please sign out and sign back in to continue.
            </p>
            <button
              onClick={this._handleSignOut}
              style={{
                background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: 8,
                padding: '10px 28px', fontSize: 14, fontWeight: 600, cursor: 'pointer',
                boxShadow: clay.shadowButton,
              }}
            >
              Sign Out & Reload
            </button>
          </div>
        </div>
      );
    }

    let content: React.ReactElement;
    // Belt-and-braces for the same module-hiding guard as _goToView: covers
    // `view` reaching render() via a path that doesn't go through
    // _goToView (persisted/restored state, a future direct setState call).
    const effectiveView: AppView = this.state.hiddenModules.indexOf(view) !== -1 ? 'dashboard' : view;
    switch (effectiveView) {
      case 'dashboard':
        content = this._renderDashboard();
        break;
      case 'list':
        content = this._renderDocumentsPage();
        break;
      case 'sites':
        content = <SitesPage host={this} />;
        break;
      case 'vessels':
        content = this._renderVesselsPage();
        break;
      // 'templates' (Templates & OCR) removed — redundant with what's
      // already on the SharePoint site. Any stale deep link / saved state
      // pointing at it falls through to the default case below (dashboard).
      case 'users':
        content = this._renderUsersPage();
        break;
      case 'settings':
        content = this._renderSettingsPage();
        break;
      case 'profile':
        content = this._renderProfilePage();
        break;
      case 'bento_email':
      case 'email_notify':
        content = this._renderBentoEmailDashboardPage();
        break;
      case 'recycle':
        content = this._renderRecycleBinPage();
        break;
      case 'archive':
        content = <ArchivePage host={this} />;
        break;
      case 'alerts':
        content = renderAlertsPage(this);
        break;
      case 'migration':
        content = this._renderMigrationAssistant();
        break;
      default:
        content = this._renderDashboard();
    }

    return (
      <>
        {this._renderLayout(content)}
        {this._renderBentoComposeModal()}
        {this._renderUploadSuccessPopup()}
        {this._renderUploadNavigationPrompt()}
        {this._renderBulkUploadModal()}
        {this._renderFileDeleteDialog()}
        {this._renderArchivePickerDialog()}
        {this._renderCreateFolderDialog()}
        {this._renderAddFolderDialog()}
        {this._renderFolderDeleteDialog()}
        {this._renderSpoDeletionToasts()}
        {renderClassifyDialog(this)}
        {renderVesselSuggestionsModal(this)}
        {/* Documents Copilot: floating button + chat panel on every page */}
        <CopilotSearchPanel host={this} />
      </>
    );
  }
}
