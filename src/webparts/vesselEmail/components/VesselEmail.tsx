import * as React from 'react';
import type { IVesselEmailProps } from './IVesselEmailProps';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from './vesselImagePool';
import { createVesselFolders, retryUntilComplete, FolderResult } from './graphFolderService';
import { MAIN_FOLDERS } from './vesselFolderTemplate';
import {
  createSyncScheduler, SyncScheduler, DeltaSyncResult,
  mergeNodeIntoMap, removeNodeFromMap, SpoFolderNode, fetchFolderChildren,
} from './deltaSync';



import type { FlatRow, GroupedRow, VesselRecord } from './types/rows';
import type { BentoEmailLog } from './types/bento';
import type { AppView, ModalMode } from './types/view';
import type {
  FormState, DocPreviewItem, DeletedNode, DocumentItem, TemplateItem,
  ApprovalItem, UserItem, FolderAnomalyItem, NormalFolderRecord, AlertItem,
  VesselSuggestion, VesselSuggestionDialog, OcrStagingItem, VesselSuggestionUploadEntry,
} from './types/ui';


import {
  cleanName, INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES,
} from './constants';
import { renderSidebar } from './pages/Sidebar';
import { renderLayout } from './pages/AppLayout';
import { renderDocPreviewDrawer } from './pages/DocPreviewDrawer';
import { renderDashboard, DashboardStats } from './pages/DashboardPage';
import { renderDocumentsPage } from './pages/DocumentsPage';
import { SitesPage } from './pages/SitesPage';
import { renderVesselsPage, renderClassifyDialog } from './pages/VesselsPage';
import { renderTemplatesPage } from './pages/TemplatesPage';
import { renderApprovalsPage } from './pages/ApprovalsPage';
import { renderReportsPage } from './pages/ReportsPage';
import { renderUsersPage } from './pages/UsersPage';
import { renderSettingsPage } from './pages/SettingsPage';
import { renderBentoEmailDashboardPage } from './pages/BentoEmailDashboardPage';
import { renderRecycleBinPage } from './pages/RecycleBinPage';
import { renderArchivePage } from './pages/ArchivePage';
import { renderAlertsPage } from './pages/AlertsPage';
import { renderBentoComposeModal } from './modals/BentoComposeModal';
import { renderVesselForm } from './modals/VesselFormModal';
import { renderDeleteModal } from './modals/DeleteVesselModal';
import { BulkUploadModal, BulkUploadFile, extractFilesFromDataTransfer } from './BulkUploadModal';
import { renderVesselSuggestionsModal } from './pages/VesselSuggestionsModal';
import { isMobileWidth, isTabletOrBelow } from './responsive';

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
          <div style={{ fontSize: 22, marginBottom: 8 }}>⚠️ Something went wrong</div>
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
              background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
              padding: '8px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
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
  loading: boolean;
  error: string | null;
  reloadKey: number;
  textFilter: string;
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
  panelLoading: boolean;
  vesselsSearch: string;
  vesselStatusFilter: string;
  vesselTypeFilter: string;

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
  documentVesselCount: number;
  documentVesselsLoadingMore: boolean;
  documentFilesLoading: boolean;
  vesselLoadingName: string | null;
  docUploadRowKey: string | null;
  docUploadBusy: boolean;
  docUploadMsg: string | null;
  folderPathStack: { id: string; name: string }[];
  uploadedFilesByFolder: Record<string, { name: string; size: string; date: string; pending?: boolean; uploading?: boolean; id?: string; uploadedAt?: number }[]>;
  selectedDocPreview: DocPreviewItem | null;
  templatesList: TemplateItem[];
  ocrStagingCount: number;
  approvalsList: ApprovalItem[];
  approvalTab: 'Pending' | 'Approved' | 'Rejected';
  // Top-header alert bell — new folder/vessel creation alerts (replaces bottom-of-module notifications)
  alertsList: AlertItem[];
  alertFilter: 'all' | 'unread';
  alertCategory: 'dms' | 'unclassified' | 'classified' | 'crud' | 'email';
  selectedAlertId: string | null;
  alertOpen: boolean;
  usersList: UserItem[];
  userSearch: string;
  userRoleFilter: string;
  reportsSelectedVessel: string;

  // Settings module state
  settingsTab: 'General' | 'Site Selection' | 'Vessel Site Provisioning' | 'Document Settings' | 'Notification Settings' | 'Permission Settings' | 'Integration' | 'Audit Logs';
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

  // Delta sync — flat id→node map representing the live SPO folder tree
  spoFolderMap: Map<string, SpoFolderNode>;
  lastDeltaSync: Date | null;

  sessionExpired: boolean;
  sessionReady: boolean;  // true once first valid session_id prop is received

  // Toast shown when a vessel is auto-moved to recycle bin via SPO deletion
  spoVesselDeletedToast: { vesselName: string; id: string } | null;
  // Toast shown when a document file/folder is moved to the SPO recycle bin
  spoDocumentDeletedToast: { itemNames: string[]; itemType: 'file' | 'folder' } | null;
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
  windowWidth: number;

  // Folder navigation history (back/forward)
  folderNavHistory: Array<{ folderPathStack: { id: string; name: string }[]; docMainFolder: State['docMainFolder'] }>;
  folderNavIndex: number;

  // File delete dialog
  fileDeleteDialog: {
    files: Array<{ id: string; name: string; folderId: string; folderPath: string }>;
    selected: Set<string>;
    busy: boolean;
    error: string | null;
  } | null;

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
  public _appUploadedFileIds: Set<string> = new Set();
  public _syncScheduler: SyncScheduler | null = null;
  private _folderRefreshSeq = 0;
  private _latestFolderRefreshTokenByKey: Map<string, number> = new Map();
  private _documentSitesLoading = false;
  private _documentLiveTreeLoading: Set<string> = new Set();
  private _documentLiveTreeLoaded: Set<string> = new Set();
  private _documentLiveTreeAbortControllers: Map<string, AbortController> = new Map();
  private _liveSharePointMerges: Set<string> = new Set();
  private _rootFoldersEnsured = false;
  public _isLoadingData = false;
  public _isUnmounted = false;
  public _deltaReloadTimer: ReturnType<typeof setTimeout> | null = null;
    public _deltaFileRefreshTimer: ReturnType<typeof setTimeout> | null = null;   // ← add this line
  public _alertRefreshTimer: ReturnType<typeof setInterval> | null = null;
  public _deleteAutoCloseTimer: ReturnType<typeof setInterval> | null = null;
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


  public constructor(props: IVesselEmailProps) {
    super(props);
    const ignoredSuggestionKeys = this._readIgnoredVesselSuggestionKeys();
    const restoredPendingSuggestions = this._readPendingVesselSuggestions().filter(
      suggestion => !ignoredSuggestionKeys.has(this._getSuggestionKey(suggestion))
    );
    const suppressUploadNavigationPrompt = this._readSuppressUploadNavigationPrompt();
    this.state = {
      rows: [],
      vessels: [],
      loading: false, error: null, reloadKey: 0, sessionExpired: Boolean(props.sessionExpired),
      textFilter: '', vesselFilter: 'all', groupFilter: 'all', catFilter: 'all', attachmentFilter: 'all',
      sort: 'default', uploadingGroupKey: null, uploadInfo: null, uploadError: null,
      selectedFileIds: new Set(), page: 0,
           modal: 'none', selectedVessel: null, deleteVesselIds: new Set(), deleteVesselProgress: {}, form: { ...BLANK_FORM },
           vesselActionPicker: null,
      modalBusy: false, modalMsg: null, modalError: null, deleteAutoCloseSeconds: null, formFieldErrors: {},
      view: 'dashboard',
      recycleBin: [], archiveList: [], panelLoading: false,
      vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all',

      documentsList: INITIAL_MOCK_DOCUMENTS,
      documentSites: [],
      documentDepartmentAliases: {},
      documentVesselAliases: {},
      tenantDiscoveredSites: [],
      activeDocumentSite: null,
      documentLiveFolders: [],
      documentLiveFoldersLoading: false,
      docViewMode: 'folder',
      docScopeType: 'vessels',
      docMainFolder: null,
      showAllVesselsInFolderView: false,
      docListPage: 0,
      docListSort: 'default',
      docGroupFilter: 'all',
      docCategoryFilter: 'all',
      docGroupLevelFilter: 'all',
      docLeafCategoryFilter: 'all',
      docSubCategoryFilter: 'all',
      documentVesselCount: 4,
      documentVesselsLoadingMore: false,
      documentFilesLoading: false,
      vesselLoadingName: null,
      docUploadRowKey: null,
      docUploadBusy: false,
      docUploadMsg: null,
      folderPathStack: [],
      uploadedFilesByFolder: {},
      selectedDocPreview: null,
      templatesList: INITIAL_MOCK_TEMPLATES,
      ocrStagingCount: 0,
      approvalsList: [],
      approvalTab: 'Pending',
      alertsList: [],
      alertFilter: 'all',
      alertCategory: 'dms',
      selectedAlertId: null,
      alertOpen: false,
      usersList: [],
      userSearch: '', userRoleFilter: 'all',
      reportsSelectedVessel: 'All Vessels',

      settingsTab: 'General',
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
      windowWidth: typeof window !== 'undefined' ? window.innerWidth : 1200,
      folderNavHistory: [{ folderPathStack: [], docMainFolder: null }],
      folderNavIndex: 0,
      uploadSuccessPopup: null,
      uploadNavigationPrompt: null,
      bulkUploadDialog: null,
      fileDeleteDialog: null,
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
      scanProgress: {
        status: 'idle', completed: 0, total: 0, title: '', recentFiles: [],
      },
    };
  }

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
        return {
          rows,
          uploadedFilesByFolder,
          docUploadMsg: `"${fileName}" archived successfully.`,
          listViewSelectedFiles: new Set(Array.from(prev.listViewSelectedFiles).filter(id => id !== fileId)),
          folderViewSelectedFiles: new Set(Array.from(prev.folderViewSelectedFiles).filter(id => id !== fileId)),
        };
      });
      const archiveData = await this._fetchJson(`${this._base()}/api/archive/nodes`);
      this.setState({ archiveList: (archiveData || []).map((item: any) => ({ ...item, name: cleanName(item.name || '') })) });
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

  public _restoreArchivedItem = async (item: DeletedNode): Promise<void> => {
    try {
      const result = await this._fetchJson(`${this._base()}/api/restore/${encodeURIComponent(item.id)}?type=${encodeURIComponent(item.kind === 'file' ? 'file' : 'folder')}&item_name=${encodeURIComponent(item.name)}`, { method: 'POST' });
      if (result?.status === 'pending') {
        this.setState({ docUploadMsg: `Restore request for "${item.name}" submitted for approval.` });
        return;
      }
      this.setState(prev => ({
        archiveList: prev.archiveList.filter(archived => archived.id !== item.id),
        docUploadMsg: `"${item.name}" restored successfully.`,
      }));
      await this._loadData(true);
    } catch (error: any) {
      this.setState({ docUploadMsg: `Restore failed: ${error?.message || 'Could not restore the item.'}` });
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
    void this._loadDashboardStats();
    void this._ensureKaizenSharePointFolders();
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

  public _loadBentoConfig(): void {
    this._fetchJson(`${this._base()}/api/email-notification/config`)
      .then((cfg: any) => {
        const recipient = cfg?.ai_bento_recipient || cfg?.graph_sender_mailbox || '';
        this.setState({ bentoConfigRecipient: recipient });
      })
      .catch(() => undefined);
  }

  public _loadDashboardStats(forceRefresh = false): Promise<void> {
    const base = this._base();
    if (!base) return Promise.resolve();
    const url = forceRefresh
      ? `${base}/api/dashboard/stats?force_refresh=true`
      : `${base}/api/dashboard/stats`;
    return this._fetchJson(url)
      .then((data: any) => {
        if (data && typeof data === 'object' && 'total_documents' in data) {
          this.setState({ dashboardStats: data as import('./pages/DashboardPage').DashboardStats });
        }
      })
      .catch(() => undefined);
  }

  public componentDidUpdate(pp: IVesselEmailProps, ps: State): void {
    // Mark session as ready once we receive a valid session_id prop
    if (!this.state.sessionReady && this.props.sessionId) {
      this.setState({ sessionReady: true });
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
    const { vesselFilter } = this.state;
    if (ps.vesselFilter !== vesselFilter && vesselFilter !== 'all') {
      this._lastRefreshedRowsKey = '';
      this._loadVesselRowsFromApi(vesselFilter).catch(() => undefined);
      setTimeout(() => this._refreshFilesFromBackendRows(), 100);
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

    // Match by folder name against known vessels
    const vessel = vessels.find(v => normName(v.name) === normName(vesselFolderNode.name));
    if (!vessel) return false;

    const now = new Date().toISOString();
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
    };

    const alert: import('./types/ui').AlertItem = {
      id: `vessel_deleted_spo_${vessel.id}_${Date.now()}`,
      drive_item_id: deletedId,
      folder_name: vessel.name,
      folder_path: vesselFolderNode.serverRelativePath,
      parent_folder_id: null,
      vessel_name: vessel.name,
      department: 'All Departments',
      created_by_email: '',
      created_by_name: 'SharePoint Online',
      alert_type: 'vessel_deleted',
      read: false,
      created_at: now,
    };

    this.setState(prev => ({
      vessels: prev.vessels.filter(v => v.id !== vessel.id),
      recycleBin: [
        recycleBinEntry,
        ...prev.recycleBin.filter(r => r.name.toLowerCase() !== vessel.name.toLowerCase()),
      ],
      alertsList: [alert, ...prev.alertsList],
      spoVesselDeletedToast: { vesselName: vessel.name, id: alert.id },
      spoDeletedVesselIds: (() => { const s = new Set(this.state.spoDeletedVesselIds); s.add(vessel.id); return s; })(),
    }));

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

    // Ensure rows are loaded for the detected vessel so the data is ready when user views alerts
    if (vesselName && !this._filesLoadedForVessels.has(vesselName)) {
      this._filesLoadedForVessels.add(vesselName);
      setTimeout(() => this._loadVesselRowsFromApi(vesselName!).catch(() => undefined), 0);
    }
  }

  public _applyDeltaResult(result: DeltaSyncResult): void {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    this.setState(prev => {
      const map = new Map(prev.spoFolderMap);

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
        removeNodeFromMap(map, id);
      }

      // Process additions/updates — detect new vessel-level folders
      const missingParentIds = new Set<string>();
      for (const node of result.added) {
        // Skip nodes we locally deleted — Graph delta may lag before reporting deletion
        if (this._appDeletedItemIds.has(node.id)) continue;

        const { missingParentId } = mergeNodeIntoMap(map, node);
        if (missingParentId) missingParentIds.add(missingParentId);

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

  public _fetchJson(url: string, optionsOrSignal?: RequestInit | AbortSignal): Promise<any> {
    const opts: RequestInit = optionsOrSignal && 'aborted' in optionsOrSignal
      ? { signal: optionsOrSignal as AbortSignal, headers: this._headers(), cache: 'no-store' }
      : {
        cache: 'no-store',
        ...(optionsOrSignal as RequestInit || {}),
        headers: {
          ...this._headers(),
          ...((optionsOrSignal as RequestInit)?.headers || {}),
        },
      };

    return fetch(url, opts)
      .then(r => {
        if (r.status === 401 || r.status === 403) {
          this.setState({ sessionExpired: true });
          throw new Error('SESSION_EXPIRED');
        }
        if (r.status >= 500) {
          throw new Error(`SERVER_ERROR_${r.status}`);
        }
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.json();
      })
      .catch(async err => {
        if (opts.signal?.aborted || err?.message === 'SESSION_EXPIRED') throw err;
        if (url.includes('nk-dms-dev.sg-nissenkaiun.com')) {
          VesselEmail._remoteServerDown = true;
          const fallbackUrl = url.replace('https://nk-dms-dev.sg-nissenkaiun.com', 'http://127.0.0.1:8000');
          console.warn(`[VesselDMS] Remote API 502/Network error (${err?.message}) — auto-switched primary base to local backend: ${fallbackUrl}`);
          try {
            const localOpts: RequestInit = {
              ...opts,
              headers: this._headersForLocalFallback(),
            };
            const r = await fetch(fallbackUrl, localOpts);
            if (r.status === 401 || r.status === 403) {
              this.setState({ sessionExpired: true });
              throw new Error('SESSION_EXPIRED');
            }
            if (!r.ok) throw new Error(`HTTP ${r.status}`);
            return await r.json();
          } catch (localErr: any) {
            if (localErr?.message === 'SESSION_EXPIRED') throw localErr;
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
    // Reload the page to force re-authentication
    window.location.reload();
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
      this.setState({ documentSites: sites, activeDocumentSite: active });
      // Fetch dynamic aliases alongside sites
      void this._loadDocumentAliases().catch(() => undefined);
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
      if (this.state.activeDocumentSite === siteKey) this.setState({ documentLiveFoldersLoading: false });
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

  public async _loadSiteDrives(siteId: string): Promise<Array<{ id: string; name: string; web_url?: string }>> {
    try {
      const data = await this._fetchJson(`${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives`);
      return Array.isArray(data?.drives) ? data.drives : [];
    } catch (err) {
      console.warn('[VesselDMS] _loadSiteDrives error:', err);
      return [];
    }
  }

  public async _loadSiteFolderChildren(
    siteId: string,
    driveId: string,
    folderId: string = 'root',
  ): Promise<{ items: any[]; summaryCounts?: any; parentPath?: string; folderId?: string }> {
    try {
      const encodedFolderRef = folderId
        .split('/')
        .map(part => encodeURIComponent(part))
        .join('/');
      const data = await this._fetchJson(
        `${this._base()}/api/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/folders/${encodedFolderRef}/children`
      );
      return {
        items: Array.isArray(data?.items) ? data.items : [],
        summaryCounts: data?.summary_counts || null,
        parentPath: data?.parent_path || '',
        folderId: data?.folder_id || folderId,
      };
    } catch (err) {
      console.warn('[VesselDMS] _loadSiteFolderChildren error:', err);
      return { items: [], summaryCounts: null, parentPath: '', folderId };
    }
  }

  public _siteDrivesCache: Map<string, Array<{ id: string; name: string; web_url?: string }>> = new Map();
  public _siteFolderItemsCache: Map<string, { items: any[]; loading?: boolean; parentPath?: string }> = new Map();
  private _siteFolderChildPrefetched: Set<string> = new Set();

  public _getOrLoadSiteDrives(siteId: string): Array<{ id: string; name: string; web_url?: string }> | null {
    if (!siteId) return [];
    if (this._siteDrivesCache.has(siteId)) {
      return this._siteDrivesCache.get(siteId) || [];
    }
    this._siteDrivesCache.set(siteId, []);
    this._loadSiteDrives(siteId).then(drives => {
      this._siteDrivesCache.set(siteId, drives);
      this.forceUpdate();
    }).catch(() => undefined);
    return null;
  }

  public _getOrLoadSiteFolderChildren(
    siteId: string,
    driveId: string,
    folderId: string = 'root',
    force: boolean = false,
  ): { items: any[]; loading?: boolean } {
    // Guard: cannot load without a valid site and drive
    if (!siteId || !driveId) {
      return { items: [], loading: false };
    }
    const key = `${siteId}::${driveId}::${folderId}`;
    if (force) {
      this._siteFolderItemsCache.delete(key);
    } else if (this._siteFolderItemsCache.has(key)) {
      return this._siteFolderItemsCache.get(key)!;
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

    const entry: { items: any[]; loading: boolean } = { items: [], loading: true };
    this._siteFolderItemsCache.set(key, entry);
    this._loadSiteFolderChildren(siteId, driveId, folderId).then(res => {
      const loadedEntry = { items: res.items || [], loading: false, parentPath: res.parentPath || '' };
      this._siteFolderItemsCache.set(key, loadedEntry);
      if (res.folderId && res.folderId !== folderId) {
        const resolvedKey = `${siteId}::${driveId}::${res.folderId}`;
        this._siteFolderItemsCache.set(resolvedKey, loadedEntry);
      }
      this.forceUpdate();
    }).catch(() => {
      this._siteFolderItemsCache.set(key, { items: [], loading: false, parentPath: '' });
      this.forceUpdate();
    });
    return entry;
  }

  public async _refreshSiteFolder(siteId: string, driveId: string, folderId: string = 'root'): Promise<void> {
    if (!siteId || !driveId) return;
    const key = `${siteId}::${driveId}::${folderId}`;
    this._siteFolderItemsCache.delete(key);
    this._getOrLoadSiteFolderChildren(siteId, driveId, folderId, true);
    this.forceUpdate();
  }

  private _usesBackendDocumentSite(): boolean {
    return false;
  }

  public async _switchDocumentSite(siteKey: string): Promise<void> {
    const base = this._base();
    const selectedSite = this.state.documentSites.find(site => site.site_key === siteKey);
    if (!selectedSite) throw new Error(`SharePoint site '${siteKey}' is not available.`);
    const response = await fetch(`${this._base()}/api/sites/active`, {
      method: 'POST',
      headers: {
        ...(base.includes('localhost') ? this._headersForLocalFallback() : this._headers()),
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ site_name: siteKey }),
    });
    if (!response.ok) throw new Error(`Could not switch site (${response.status})`);
    this._siteFolderItemsCache.clear();
    this._siteFolderChildPrefetched.clear();
    this.setState({
      activeDocumentSite: siteKey,
      documentLiveFolders: [],
      documentLiveFoldersLoading: false,
      rows: [],
      uploadedFilesByFolder: {},
      folderPathStack: [],
      docMainFolder: null,
      folderNavHistory: [{ folderPathStack: [], docMainFolder: null }],
      folderNavIndex: 0,
    });
    await this._loadData(true);
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
            return { alertsList: [...localOnly, ...(res as AlertItem[])] };
          });
        }
      })
      .catch(() => undefined);
  };

  /** Add a local activity alert and immediate recycle-bin popup for a deleted SPO item. */
  public _handleSpoDocumentDeletion(node: import('./deltaSync').SpoFolderNode): void {
    const now = new Date().toISOString();
    const itemType = node.isFolder ? 'folder' : 'file';
    const alert: AlertItem = {
      id: `document_deleted_${node.id}_${Date.now()}`,
      drive_item_id: node.id,
      folder_name: node.name,
      folder_path: node.serverRelativePath || 'SharePoint Online Documents',
      parent_folder_id: node.parentId,
      vessel_name: null,
      department: 'Documents',
      created_by_email: this.props.userEmail || '',
      created_by_name: this.props.userDisplayName || 'SharePoint Online',
      alert_type: 'document_deleted',
      alert_category: 'crud',
      read: false,
      created_at: now,
      item_type: itemType,
      spo_path: node.serverRelativePath,
    };

    this.setState(prev => {
      const alreadyTracked = prev.alertsList.some(
        a => a.alert_type === 'document_deleted' && a.drive_item_id === node.id,
      );
      if (alreadyTracked) return null;
      return {
        alertsList: [alert, ...prev.alertsList],
        spoDocumentDeletedToast: {
          itemNames: [...(prev.spoDocumentDeletedToast?.itemNames || []), node.name],
          itemType,
        },
      };
    });
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
          this._handleSpoDocumentDeletion(node);
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

  public _toggleAlertBell = (): void => {
    this.setState(prev => ({ alertOpen: !prev.alertOpen }));
    if (!this.state.alertOpen) {
      this._fetchAlerts();
    }
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

  public _setAlertCategory = (category: 'dms' | 'unclassified' | 'classified' | 'crud' | 'email'): void => {
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
      try {
        vesselList = await this._fetchJson(`${base}/api/vessels`, signal);
        vesselListFetched = true;
      } catch {
        /* ignore */
      }
      if (signal.aborted) return;

      let vessels: VesselRecord[] = [];
      const { spoDeletedVesselIds, recycleBin } = this.state;
      const deletedNames = new Set(
        recycleBin
          .filter(r => r.kind === 'vessel' || r.item_type === 'vessel')
          .map(r => r.name.toLowerCase())
      );
      if (vesselList && Array.isArray(vesselList) && vesselList.length > 0) {
        vessels = vesselList
          .map((v: any) => ({ ...v, name: cleanName(v.name), status: v.status || 'Active' }))
          .filter((v: any) => !spoDeletedVesselIds.has(v.id) && !deletedNames.has((v.name || '').toLowerCase()));
      }
      // Preserve any locally created vessels currently in state that might not yet be returned by backend
      const fetchedNames = new Set(vessels.map(v => (v.name || '').trim().toLowerCase()));
      const pendingLocalVessels = (this.state.vessels || []).filter(v =>
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

            // 2. Scan Technical & Crewing level
            const spoVesselNodes = await this._getGraphChildren('Technical & Crewing', signal).catch(() => []);
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
          this._refreshFilesFromBackendRows(0);
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
        this._lastRefreshedRowsKey = '';
        this._refreshFilesFromBackendRows();
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
  public async _openSharePointFolder(row: GroupedRow): Promise<void> {
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
      const paths = [
        liveFolderPath,
        sharePointPath,
        sharePointPath.replace(/\/Vessels\//i, '/'),
        `Vessels/${sharePointPath}`,
        `Vessels/Specific Vessels/${sharePointPath}`,
      ];
      let lastError: any;
      for (const candidatePath of Array.from(new Set(paths.filter(Boolean))) as string[]) {
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
  ): Promise<{ fileId: string | null; statusPending: boolean; folderId: string | null; isGraphUpload: boolean }> {
    const base = this._base();
    if (!base) return { fileId: null, statusPending: false, folderId: null, isGraphUpload: false };

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

        // 3. Upload file directly into the resolved SharePoint Online folder
        let item: any = null;
        if (file.size <= 4 * 1024 * 1024) {
          const uploadUrl = `/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}:/${encodeURIComponent(file.name)}:/content`;
          item = await graphClient.api(uploadUrl).put(file);
        } else {
          // Large file chunked upload session (> 4 MB)
          const sessionUrl = `/sites/${effectiveSiteId}/drives/${effectiveDriveId}/items/${folder.id}:/${encodeURIComponent(file.name)}:/createUploadSession`;
          const session = await graphClient.api(sessionUrl).post({
            item: { '@microsoft.graph.conflictBehavior': 'rename', name: file.name },
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

        return { fileId: item.id as string, statusPending: false, folderId: folder.id as string, isGraphUpload: true };
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
    return { fileId: data?.id || null, statusPending: data?.status === 'pending', folderId: restFolderId || null, isGraphUpload: false };
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

    // Walk each main folder in parallel
    await this._mapLimit(
      this.MAIN_FOLDER_NAMES,
      3,
      async mainFolderName => {
        if (signal.aborted) return;
        let vesselPath = `${this.VESSEL_ROOT}/${vesselName}/${mainFolderName}`;
        // Get top-level categories inside this vessel's main folder
        let topCats = await this._getGraphChildren(vesselPath, signal).catch(() => []);
        // Keep existing pre-migration vessel folders readable.
        if (topCats.length === 0) {
          const legacyPaths = [
            `${mainFolderName}/Vessels/${vesselName}`,
            `${mainFolderName}/${vesselName}`,
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

  /**
   * Ensures the complete Kaizen - Knowledge Bank folder hierarchy exists
   * directly in the active SharePoint Online document library.
   */
  public async _ensureKaizenSharePointFolders(): Promise<void> {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    const ensureFolderInDrive = async (parentPath: string, folderName: string): Promise<string | null> => {
      const cleanParent = (parentPath || '').replace(/^\/+|\/+$/g, '');
      const fullPath = cleanParent ? `${cleanParent}/${folderName}` : folderName;
      const encodedFullPath = fullPath.split('/').map(s => encodeURIComponent(s)).join('/');
      try {
        const existing = await graphClient
          .api(`/sites/${siteId}/drives/${driveId}/root:/${encodedFullPath}?$select=id,name,folder`)
          .get();
        if (existing?.id) return existing.id;
      } catch {
        // Not found, proceed to create
      }

      const createUrl = cleanParent
        ? `/sites/${siteId}/drives/${driveId}/root:/${cleanParent.split('/').map(s => encodeURIComponent(s)).join('/')}:/children`
        : `/sites/${siteId}/drives/${driveId}/root/children`;

      try {
        const created = await graphClient.api(createUrl).post({
          name: folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'fail',
        });
        return created?.id || null;
      } catch {
        try {
          const fallback = await graphClient
            .api(`/sites/${siteId}/drives/${driveId}/root:/${encodedFullPath}?$select=id,name,folder`)
            .get();
          return fallback?.id || null;
        } catch {
          return null;
        }
      }
    };

    try {
      // 1. Root folder
      await ensureFolderInDrive('', 'Kaizen - Knowledge Bank');

      // 2. Sections
      await ensureFolderInDrive('Kaizen - Knowledge Bank', 'Templates');
      await ensureFolderInDrive('Kaizen - Knowledge Bank', 'Procedures and Work Instructions');
      await ensureFolderInDrive('Kaizen - Knowledge Bank', 'Lessons Learned');
      await ensureFolderInDrive('Kaizen - Knowledge Bank', 'Circulars and Guidance');

      // 3. Sub-categories under Circulars and Guidance
      const circularsSubs = [
        'Equipment Maker',
        'Class',
        'Flag - Port State',
        'SIRE-OCIMF-RightShip',
        'Shipyard',
      ];
      for (const sub of circularsSubs) {
        await ensureFolderInDrive('Kaizen - Knowledge Bank/Circulars and Guidance', sub);
      }
      console.log('[VesselDMS] _ensureKaizenSharePointFolders: verified and created Kaizen folders in SharePoint Online');
    } catch (e) {
      console.warn('[VesselDMS] _ensureKaizenSharePointFolders error:', e);
    }
  }

  // ── Navigation & Views ────────────────────────────────────────────────────

  /** Push a new folder navigation entry, truncating any forward history. */
  public _pushFolderNav(
    folderPathStack: { id: string; name: string }[],
    docMainFolder: State['docMainFolder'],
  ): void {
    if (
      docMainFolder === 'Kaizen - Knowledge Bank' ||
      (folderPathStack.length > 0 && (folderPathStack[0]?.name === 'Kaizen - Knowledge Bank' || folderPathStack[0]?.id === 'kaizen_root'))
    ) {
      void this._ensureKaizenSharePointFolders();
    }
    this.setState(prev => {
      const truncated = prev.folderNavHistory.slice(0, prev.folderNavIndex + 1);
      const next = [...truncated, { folderPathStack, docMainFolder }];
      return { folderNavHistory: next, folderNavIndex: next.length - 1, folderPathStack, docMainFolder };
    });
  }

  public _goToView = async (view: AppView): Promise<void> => {
    this.setState({ view });
    if (view === 'dashboard') {
      void this._loadDashboardStats();
    } else if (view === 'vessels') {
      this.setState({ panelLoading: true });
      try {
        const data = await this._fetchJson(`${this._base()}/api/vessels`);
        if (data && Array.isArray(data)) {
          const { spoDeletedVesselIds, recycleBin } = this.state;
          // Filter out vessels that were soft-deleted via SPO delta sync
          // (backend DELETE may still be in-flight or the DB may not have updated yet)
          const deletedNames = new Set(
            recycleBin
              .filter(r => r.kind === 'vessel' || r.item_type === 'vessel')
              .map(r => r.name.toLowerCase())
          );
          this.setState({
            vessels: data
              .map((v: any) => ({ ...v, name: cleanName(v.name), status: v.status || 'Active' }))
              .filter((v: any) => !spoDeletedVesselIds.has(v.id) && !deletedNames.has((v.name || '').toLowerCase())),
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
      this.setState({ panelLoading: true, archiveList: [] });
      try {
        const data = await this._fetchJson(`${this._base()}/api/archive/nodes`);
        this.setState({ archiveList: (data || []).map((v: any) => ({ ...v, name: cleanName(v.name || '') })), panelLoading: false });
      } catch {
        this.setState({ archiveList: [], panelLoading: false });
      }
    } else if (view === 'approvals') {
      this.setState({ panelLoading: true });
      try {
        const userEmail = this.props.userEmail || '';
        const base = this._base();

        const myDataPromise = userEmail
          ? this._fetchJson(`${base}/api/my-approvals`).catch(() => null)
          : Promise.resolve(null);

        const adminDataPromise = userEmail
          ? this._fetchJson(`${base}/api/approvals?admin=${encodeURIComponent(userEmail)}`).catch(() => null)
          : Promise.resolve(null);

        const [myData, adminData] = await Promise.all([myDataPromise, adminDataPromise]);

        const combined: any[] = [];
        const seen = new Set<string>();
        for (const a of [
          ...(Array.isArray(adminData) ? adminData : []),
          ...(Array.isArray(myData) ? myData : []),
        ]) {
          const key = String(a.id || a._id || '');
          if (!seen.has(key)) { seen.add(key); combined.push(a); }
        }

        const mapDate = (v: any): string => {
          const raw = v?.created_at || v?.uploaded_at || v?.requestedOn || v?.requested_on;
          if (!raw) return '—';
          try { return new Date(raw).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); } catch { return raw; }
        };

        const mapped: ApprovalItem[] = combined
          // Strip out any 'cancelled' or 'completed' entries (activity logs) that slip through
          .filter((a: any) => !['cancelled', 'completed'].includes((a.status || '').toLowerCase()))
          .map((a: any) => ({
          id: String(a.id || a._id || Date.now()),
          documentName: a.document_name || a.filename || a.file_name || a.fileName || a.target_description || a.message || a.document_name || (a.action_type ? a.action_type.replace(/_/g, ' ') : 'Unknown'),
          vessel: a.vessel_name || a.vesselName || a.vessel || '—',
          requestedBy: a.uploaded_by_email || a.uploaded_by_name || a.requested_by || a.requestedBy || a.uploader_email || a.uploader || a.user_email || '—',
          requestedOn: mapDate(a),
          status: a.status === 'approved' ? 'Approved' : a.status === 'rejected' ? 'Rejected' : 'Pending',
          actionType: a.action_type || 'upload',
        }));
        // Merge: keep any locally-added pending items not yet in the backend response
        const backendIds = new Set(mapped.map(m => m.id));
        const localOnly = this.state.approvalsList.filter(a => !backendIds.has(a.id));
        this.setState({ approvalsList: [...mapped, ...localOnly], panelLoading: false });
      } catch {
        this.setState({ panelLoading: false });
      }
    } else if (view === 'users') {
      this.setState({ panelLoading: true });
      try {
        const data = await this._fetchJson(`${this._base()}/api/users`);
        const asRole = (value: any): UserItem['role'] => {
          const role = String(value || '').toLowerCase();
          if (role === 'administrator') return 'Administrator';
          if (role === 'manager') return 'Manager';
          if (role === 'reviewer') return 'Reviewer';
          return 'User';
        };
        const asStatus = (value: any): UserItem['status'] => {
          return String(value || '').toLowerCase() === 'active' ? 'Active' : 'Inactive';
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

  // ── CRUD Handlers ─────────────────────────────────────────────────────────

   public _openCreate = (): void => {
    this.setState({ modal: 'create', selectedVessel: null, form: { ...BLANK_FORM }, modalMsg: null, modalError: null, formFieldErrors: {} });
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

  public _closeModal = (): void => {
    this._clearDeleteAutoCloseTimer();
    if (!this.state.modalBusy) this.setState({ modal: 'none', vesselActionPicker: null, modalMsg: null, modalError: null, deleteVesselProgress: {}, deleteAutoCloseSeconds: null });
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
    this.setState({ modalBusy: true, modalError: null, modalMsg: null });

    const newVesselRecord: VesselRecord = {
      id: `v_${Date.now()}`,
      name: form.name.trim(),
      imo: form.imo.trim(),
      shipyard: form.shipyard.trim() || undefined,
      hull_number: form.hull_number.trim() || undefined,
      vessel_type: form.vessel_type || undefined,
      status: 'Active',
      image_url: pickRandomVesselImage(form.vessel_type),
      // Include target site IDs so link generation works immediately before backend refresh
      provisioned_site_ids: form.target_site_ids && form.target_site_ids.length > 0 ? form.target_site_ids : undefined,
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
          provisioned_site_ids: form.target_site_ids && form.target_site_ids.length > 0 ? form.target_site_ids : undefined,
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

      // Transition modal to success screen immediately
      this.setState({
        modalBusy: false,
        modalMsg: `🎉 Vessel "${form.name}" Created & Provisioned Successfully!`,
        modalError: null,
        formFieldErrors: {},
      });
    } catch (e: any) {
      // A real network failure (fetch itself threw, server unreachable) —
      // NOT a validation rejection, that's handled above and returns early.
      this.setState(prev => ({
        vessels: [newVesselRecord, ...prev.vessels.filter(v => v.name.toLowerCase() !== newVesselRecord.name.toLowerCase())],
        modalBusy: false,
        modalMsg: `🎉 Vessel "${form.name}" Created & Provisioned Successfully!`,
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
    const { deleteVesselIds, vessels } = this.state;
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
        const res = await fetch(`${this._base()}/api/vessels/${vessel.id}?vessel_name=${encodeURIComponent(vessel.name)}`, {
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
    const immediateRecycleBinEntries: DeletedNode[] = selected
      .filter(v => deletedIds.includes(v.id))
      .map(v => ({
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
      }));

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
          rows: [],
        });
        await this._loadData(true);
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
          rows: [],
        });
        await this._loadData(true);
        return { success: true, results: [] };
      } catch (error: any) {
        const message = error?.message || 'Folder provisioning failed.';
        this.setState({ folderCreationBusy: false, folderProvisioningVesselId: null, folderCreationResults: null, folderCreationError: message });
        return { success: false, results: [] };
      }
    }

    const { graphClient, siteId, driveId } = this.props;
    console.log(`[VesselDMS] _provisionVesselFolders vessel="${vesselName}" graphClient=${!!graphClient} siteId="${siteId}" driveId="${driveId}"`);
    if (!graphClient || !siteId || !driveId) {
      console.error('[VesselDMS] _provisionVesselFolders aborted — missing graphClient/siteId/driveId. Props:', { graphClient: !!graphClient, siteId, driveId });
      this.setState({ folderCreationError: 'SharePoint context not ready. Please refresh the page and try again.' });
      return { success: false, results: [] };
    }

        this.setState({ folderCreationBusy: true, folderCreationResults: [], folderCreationError: null, folderProvisioningVesselId: vesselId || null });
    try {
      const OVERALL_TIMEOUT_MS = 4 * 60 * 1000;
      const onProgress = (progressResult: FolderResult): void => {
        this.setState(prev => ({ folderCreationResults: [...(prev.folderCreationResults || []), progressResult] }));
      };
      const result = await Promise.race([
        createVesselFolders(graphClient, siteId, driveId, vesselName, onProgress, this._rootFoldersEnsured),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error(
            'Folder provisioning is taking much longer than expected and may have stalled. ' +
            'Please close this dialog and try again, or check SharePoint directly to see what was created so far.',
          )), OVERALL_TIMEOUT_MS),
        ),
      ]);
      this._rootFoldersEnsured = true;

      // ── Auto-retry any failed folders ───────────────────────────────────
      // If some folders failed on the first pass (throttling / transient Graph
      // errors), automatically retry them up to 3 more times (3 s between
      // rounds) before giving up.  Each round updates folderCreationResults so
      // the live ticker in the dialog reflects current status.
      let finalResults = result.results;
      const initialFailCount = finalResults.filter(r => r.status === 'failed').length;
      if (initialFailCount > 0) {
        console.log(`[VesselDMS] ${initialFailCount} folder(s) failed — starting auto-retry loop`);
        finalResults = await retryUntilComplete(
          graphClient,
          siteId,
          driveId,
          finalResults,
          /* maxAttempts */ 3,
          /* delayMs     */ 3000,
          (_attempt, _total, updated) => {
            // Push updated results to state so the ticker reflects each round
            this.setState({ folderCreationResults: updated });
          },
        );
      }
      // ────────────────────────────────────────────────────────────────────

      const allSuccess = finalResults.every(r => r.status !== 'failed');
      const provisionedSet = new Set(this.state.provisionedVesselIds);
      if (vesselId) {
        provisionedSet.add(vesselId);
        // Persist provisioned status in database
        void fetch(`${this._base()}/api/vessels/${vesselId}/provision`, {
          method: 'POST',
          headers: this._headers(),
        }).catch(() => undefined);
      }
      this.setState(prev => {
        const nextState: any = {
          folderCreationBusy: false,
          folderProvisioningVesselId: null,
          folderCreationResults: finalResults,
          folderCreationError: allSuccess ? null : 'Some folders could not be created after retrying. Check the creation log.',
          provisionedVesselIds: provisionedSet,
        };
        // If the provision dialog is currently open for this vessel, transition it to DONE state
        if (prev.spoProvisionDialog) {
          const failedCount = finalResults.filter(r => r.status === 'failed').length;
          nextState.spoProvisionDialog = {
            ...prev.spoProvisionDialog,
            provisioning: false,
            done: true,
            error: allSuccess ? null : `${failedCount} folder${failedCount === 1 ? '' : 's'} could not be created — see details below.`,
          };
        }
        return nextState;
      });
      // Trigger an immediate delta sync so the new folders appear in the tree
      void this._syncScheduler?.triggerNow().catch(() => undefined);
      return { success: allSuccess, results: finalResults };
    } catch (err: any) {
      const msg = err?.message ?? 'Folder provisioning failed.';
      this.setState(prev => {
        const nextState: any = {
          folderCreationBusy: false,
          folderProvisioningVesselId: null,
          folderCreationError: msg,
        };
        if (prev.spoProvisionDialog) {
          nextState.spoProvisionDialog = {
            ...prev.spoProvisionDialog,
            provisioning: false,
            done: true,
            error: msg,
          };
        }
        return nextState;
      });
      return { success: false, results: [] };
    }
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

  /** Share one Graph children request across overlapping refreshes. */
  private async _getGraphChildrenResponse(url: string): Promise<any> {
    const cached = this._graphChildrenCache.get(url);
    if (cached && cached.expiresAt > Date.now()) return cached.value;
    const active = this._graphChildrenInFlight.get(url);
    if (active) return active;
    const request = (async () => {
      const result: any = await this.props.graphClient!.api(url).get();
      this._graphChildrenCache.set(url, { value: result, expiresAt: Date.now() + 5000 });
      return result;
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
    files: Array<{ id: string; name: string; folderId: string; folderPath: string }>
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
        // Graph DELETE on a drive item moves it to SharePoint's own Recycle Bin (soft delete).
        // Only use Graph for real drive item IDs — numeric-only IDs are backend DB IDs.
        if (graphClient && siteId && driveId && file.id && !/^file_/.test(file.id) && !/^\d+$/.test(file.id)) {
          // DELETE /drives/{driveId}/items/{itemId} → soft delete (moves to SPO Recycle Bin)
          await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${file.id}`).delete();
        } else {
          // Fallback: backend soft-delete
          const res = await fetch(`${this._base()}/api/files/${encodeURIComponent(file.id)}`, {
            method: 'DELETE', headers: this._headers(),
          });
          if (!res.ok && res.status !== 202) throw new Error(`HTTP ${res.status}`);
        }
        deletedIds.push(file.id);

        const usedGraph = graphClient && siteId && driveId && file.id && !/^file_/.test(file.id) && !/^\d+$/.test(file.id);
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
        };
        if (usedGraph) this._appDeletedItemIds.add(file.id);
        this._handleSpoDocumentDeletion({
          id: file.id,
          name: file.name,
          parentId: null,
          isFolder: false,
          serverRelativePath: file.folderPath,
          children: [],
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
          background: '#fff', borderRadius: 16, padding: isMobile ? '16px 14px' : '28px 32px', width: isMobile ? '95vw' : 'auto', minWidth: isMobile ? 0 : 420, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto',
          boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Segoe UI', sans-serif",
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>🗑</div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: '#0f172a' }}>Delete Files</div>
              <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Check files to move to Recycle Bin</div>
            </div>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 12, maxHeight: 260, overflowY: 'auto' }}>
            {files.map(f => (
              <label key={f.id} style={{
                display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px',
                borderRadius: 8, border: `1.5px solid ${selected.has(f.id) ? '#ef4444' : '#e2e8f0'}`,
                background: selected.has(f.id) ? '#fff5f5' : '#f8fafc', cursor: 'pointer',
              }}>
                <input
                  type="checkbox"
                  checked={selected.has(f.id)}
                  onChange={() => toggle(f.id)}
                  style={{ width: 16, height: 16, accentColor: '#ef4444', cursor: 'pointer' }}
                />
                <span style={{ fontSize: 16 }}>📄</span>
                <span style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', flex: 1, wordBreak: 'break-all' }}>{f.name}</span>
              </label>
            ))}
          </div>

          {/* Additional files from the same folder that can be added */}
          {additionalFiles.length > 0 && (
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', marginBottom: 6 }}>
                + Add more files from this folder
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 140, overflowY: 'auto' }}>
                {additionalFiles.map(f => (
                  <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', borderRadius: 8, border: '1px dashed #cbd5e1', background: '#f8fafc' }}>
                    <span style={{ fontSize: 14 }}>📄</span>
                    <span style={{ fontSize: 12, color: '#334155', flex: 1, wordBreak: 'break-all' }}>{f.name}</span>
                    <button
                      onClick={() => addFileToDialog(f)}
                      style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#334155', whiteSpace: 'nowrap' }}
                    >
                      + Add
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {error && (
            <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#dc2626', marginBottom: 12 }}>
              ⚠️ {error}
            </div>
          )}

          <div style={{ background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#c2410c', marginBottom: 20 }}>
            ⚠️ {selected.size} file{selected.size !== 1 ? 's' : ''} will be moved to the Recycle Bin. This can be undone from the Recycle Bin page.
          </div>

          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', flexDirection: isMobile ? 'column' : 'row' }}>
            <button
              onClick={() => this.setState({ fileDeleteDialog: null })}
              disabled={busy}
              style={{ minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: '1px solid #cbd5e1', background: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', color: '#334155' }}
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
              {busy ? '⏳ Deleting…' : `🗑 Delete ${selected.size} file${selected.size !== 1 ? 's' : ''}`}
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
    try {
      if (backendFolderKey && !backendFolderKey.startsWith('sf_')) {
        await fetch(`${this._base()}/api/folders/${encodeURIComponent(backendFolderKey)}?folder_name=${encodeURIComponent(target.name)}`, {
          method: 'DELETE', headers: this._headers(),
        }).catch(() => undefined);
      }
    } catch { /* non-critical */ }

    // 3. Create Recycle Bin entry
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
    };

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
          background: '#fff', borderRadius: 16, width: isMobile ? '95vw' : 480, maxWidth: '95vw', maxHeight: '90vh', overflow: 'hidden',
          boxShadow: '0 20px 60px rgba(15,23,42,0.3)', border: '1px solid #fee2e2',
        }}>
          {/* Header */}
          <div style={{ padding: '20px 24px 16px', background: '#fff1f2', borderBottom: '1px solid #fecdd3' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <div style={{ width: 42, height: 42, borderRadius: 10, background: '#ffe4e6', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>
                🗑️
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
                ⚠️ {error}
              </div>
            )}

            <div style={{ fontSize: 13, color: '#334155', fontWeight: 600, marginBottom: 8 }}>
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
                      borderRadius: 10, border: `1.5px solid ${isChecked ? '#e11d48' : '#e2e8f0'}`,
                      background: isChecked ? '#fff1f2' : '#f8fafc', cursor: 'pointer',
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
                    <span style={{ fontSize: 20 }}>📁</span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a' }}>
                        {folder.name} {folder.isCurrent ? <span style={{ fontSize: 11, color: '#e11d48', fontWeight: 600 }}>(Current Folder)</span> : ''}
                      </div>
                      <div style={{ fontSize: 11, color: '#64748b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={folder.path}>
                        {folder.path}
                      </div>
                    </div>
                  </label>
                );
              })}
            </div>

            <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: 12, fontSize: 12, color: '#64748b', lineHeight: 1.5 }}>
              ℹ️ Moving <strong style={{ color: '#0f172a' }}>{selectedFolder?.name}</strong> to the Recycle Bin will also delete its contained files in SharePoint Online. You can restore this folder anytime from the <strong style={{ color: '#0f172a' }}>Recycle Bin</strong> page.
            </div>
          </div>

          {/* Footer */}
          <div style={{ padding: '14px 24px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', justifyContent: 'flex-end', gap: 10, flexDirection: isMobile ? 'column' : 'row' }}>
            <button
              onClick={() => this.setState({ folderDeleteDialog: null })}
              disabled={busy}
              style={{
                border: '1px solid #cbd5e1', background: '#fff', borderRadius: 8,
                minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 18px', fontSize: 13, fontWeight: 600, color: '#475569',
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
              {busy ? 'Moving to Recycle Bin…' : '🗑 Move to Recycle Bin'}
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
        background: '#fff', borderRadius: 16, boxShadow: '0 8px 40px rgba(0,0,0,0.18)',
        border: '1.5px solid #86efac', padding: isMobile ? '12px' : '20px 24px 18px', width: isMobile ? 'calc(100vw - 20px)' : 'auto', minWidth: isMobile ? 0 : 340, maxWidth: isMobile ? 'calc(100vw - 20px)' : 420,
        fontFamily: "'Segoe UI', sans-serif", animation: 'slideInRight 0.3s ease',
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
              {p.isPending ? '⏳' : '✅'}
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>
                {p.isPending ? 'Submitted for Approval' : 'Upload Successful!'}
              </div>
              <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>
                Auto-closes in {p.secondsLeft}s
              </div>
            </div>
          </div>
          <button
            onClick={() => { clearInterval(this._uploadSuccessTimer!); this.setState({ uploadSuccessPopup: null }); }}
            style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#94a3b8', minHeight: 44, minWidth: 44, fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}
            title="Close"
          >✕</button>
        </div>
        {/* File info */}
        <div style={{ background: '#f8fafc', borderRadius: 8, padding: '8px 12px', marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: '#64748b', marginBottom: 3 }}>📄 File</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', wordBreak: 'break-all' }}>{p.fileName}</div>
          <div style={{ fontSize: 11, color: '#64748b', marginTop: 6 }}>📁 Path</div>
          <div style={{ fontSize: 12, color: '#334155', marginTop: 2, wordBreak: 'break-all' }}>{p.destinationPath}</div>
          {Array.isArray(p.unidentifiedFiles) && p.unidentifiedFiles.length > 0 && (
            <div style={{ marginTop: 10, background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 8, padding: '8px 10px' }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: '#9a3412' }}>⚠ Vessel name not identified files</div>
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
              this.setState({ uploadSuccessPopup: null, alertOpen: true });
              this._fetchAlerts();
            }}
            style={{
              flex: 1, background: '#0078d4', color: '#fff', border: 'none', borderRadius: 8,
              minHeight: 44, padding: '8px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer',
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
                background: '#0078d4', color: '#fff', borderRadius: 8,
                minHeight: 44, padding: '8px 12px', fontSize: 12, fontWeight: 600, textDecoration: 'none',
                cursor: 'pointer',
              }}
            >
              🔗 Open in SharePoint
            </a>
          )}
          <button
            onClick={() => { clearInterval(this._uploadSuccessTimer!); this.setState({ uploadSuccessPopup: null }); }}
            style={{
              flex: 1, background: '#f1f5f9', color: '#334155', border: 'none', borderRadius: 8,
              minHeight: 44, padding: '8px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            Stay here
          </button>
        </div>
        {/* Countdown progress bar */}
        <div style={{ marginTop: 12, height: 3, background: '#f1f5f9', borderRadius: 2, overflow: 'hidden' }}>
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
      void this._provisionVesselFolders(newVessel.name, newVessel.id).catch(() => undefined);
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

  public _renderTemplatesPage(): React.ReactElement {
    return renderTemplatesPage(this);
  }

  public _renderApprovalsPage(): React.ReactElement {
    return renderApprovalsPage(this);
  }

  public _renderReportsPage(): React.ReactElement {
    return renderReportsPage(this);
  }

  public _renderUsersPage(): React.ReactElement {
    return renderUsersPage(this);
  }

  public _renderSettingsPage(): React.ReactElement {
    return renderSettingsPage(this);
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
      const base = this._base();
      const names: string[] = [];

      // 1. Trigger on-demand loading of files for this vessel if not already loaded
      if (!this._filesLoadedForVessels.has(vesselName)) {
        this._filesLoadedForVessels.add(vesselName);
        await this._mergeLiveSharePointFiles([vesselName]).catch(() => undefined);
      }

      // 2. Fetch from /api/my-approvals?status=approved (or /api/approvals?status=approved)
      const approvals = await this._fetchJson(`${base}/api/my-approvals?status=approved`).catch(() =>
        this._fetchJson(`${base}/api/approvals?status=approved`).catch(() => [] as any[])
      );
      const fileIdUpdates: Record<string, string> = {};
      if (Array.isArray(approvals)) {
        approvals.forEach((a: any) => {
          const aVessel = (a.vessel_name || a.vesselName || '').trim().toLowerCase();
          const fname = a.file_name || a.fileName || a.filename || a.name;
          if (fname && (aVessel === vLower || !aVessel) && names.indexOf(fname) === -1) {
            names.push(fname);
          }
          const realFileId = a.drive_item_id || a.driveItemId || a.file_id || a.fileId;
          if (fname && realFileId) {
            fileIdUpdates[`${vLower}||${fname}`] = realFileId;
          }
        });
      }

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


  public _renderBentoComposeModal(): React.ReactElement | null {
    return renderBentoComposeModal(this);
  }


  // ── Existing Views: Recycle Bin ───────────────────────────────────────────

  /**
   * Load all deleted items for the Recycle Bin page:
   * 1. Fetches backend deleted nodes database records.
   * 2. Queries SharePoint Online Recycle Bin (web and site collection) to find deleted folders and files.
   * 3. Merges and deduplicates with active in-memory session deletions.
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
      if (this.props.siteUrl) {
        const endpoints = [
          `${this.props.siteUrl}/_api/web/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType,DeletedDate,DeletedByName&$top=500`,
          `${this.props.siteUrl}/_api/site/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType,DeletedDate,DeletedByName&$top=500`,
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
                });
              }
            }
          } catch (err) {
            console.warn('[VesselDMS] Fetching SPO RecycleBin items warning:', err);
          }
        }
      }

      this.setState(prev => {
        const prevLocal = prev.recycleBin || [];
        const combined = [...backendItems];

        // Add SPO recycled nodes not already present in combined
        for (const spoNode of spoRecycledNodes) {
          const exists = combined.some(b =>
            b.id === spoNode.id ||
            b.recycle_bin_item_id === spoNode.id ||
            (b.name.toLowerCase() === spoNode.name.toLowerCase() && b.original_path === spoNode.original_path)
          );
          if (!exists) combined.push(spoNode);
        }

        // Add locally tracked items not in combined
        for (const locNode of prevLocal) {
          const exists = combined.some(c =>
            c.id === locNode.id ||
            (c.recycle_bin_item_id && locNode.recycle_bin_item_id && c.recycle_bin_item_id === locNode.recycle_bin_item_id) ||
            (c.name.toLowerCase() === locNode.name.toLowerCase() && c.original_path === locNode.original_path)
          );
          if (!exists) combined.push(locNode);
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

    try {
      // 1. Get CSRF digest
      let digest = '';
      try {
        const digestRes = await fetch(`${this.props.siteUrl}/_api/contextinfo`, {
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
        `${this.props.siteUrl}/_api/web/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType&$top=500`,
        `${this.props.siteUrl}/_api/site/RecycleBin?$select=Id,LeafName,Title,DirName,ItemType&$top=500`,
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
          `${this.props.siteUrl}/_api/web/RecycleBin(guid'${guid}')/restore()`,
          `${this.props.siteUrl}/_api/site/RecycleBin(guid'${guid}')/restore()`,
          `${this.props.siteUrl}/_api/web/RecycleBin('${guid}')/restore()`,
          `${this.props.siteUrl}/_api/site/RecycleBin('${guid}')/restore()`,
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
        if (!restoredThis && this.props.graphClient && this.props.siteId) {
          try {
            await this.props.graphClient
              .api(`/sites/${this.props.siteId}/recycleBin/items/${guid}/restore`)
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

      // 5. If it's a vessel, also ensure the folder structure is intact via provisioning fallback
      if (isVessel && this.props.graphClient && this.props.siteId && this.props.driveId) {
        try {
          void this._provisionVesselFolders(item.name).catch(() => undefined);
        } catch {
          // Ignore
        }
      }

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

    try {
      if (graphClient && siteId && driveId && isRealGraphId) {
        try {
          // Graph /recycleBin is not supported in v1.0 - use SharePoint REST API instead.
          const findInRecycleBin = async (): Promise<string | null> => {
            if (!this.props.siteUrl) return null;
            try {
              const nameLower = item.name.toLowerCase();
              const rbRes = await fetch(
                `${this.props.siteUrl}/_api/site/RecycleBin?$filter=LeafName eq '${encodeURIComponent(item.name)}'&$select=Id,LeafName&$top=10`,
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
            const r = await fetch(`${this.props.siteUrl}/_api/contextinfo`, {
              method: 'POST',
              headers: { Accept: 'application/json;odata=nometadata' },
            });
            const d = await r.json();
            return d?.FormDigestValue ?? d?.['odata.metadata'] ?? '';
          };
          const deleteFromRecycleBin = async (rbId: string): Promise<void> => {
            if (!this.props.siteUrl) return;
            const digest = await getSpoDigest();
            const res = await fetch(
              `${this.props.siteUrl}/_api/site/RecycleBin('${rbId}')`,
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

  public _renderArchivePage(): React.ReactElement {
    return renderArchivePage(this);
  }

  public _renderSpoVesselDeletedToast(): React.ReactElement | null {
    const { spoVesselDeletedToast } = this.state;
    if (!spoVesselDeletedToast) return null;
    const isMobile = isMobileWidth(this.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
    return (
      <div style={{
        position: 'fixed', bottom: isMobile ? 10 : 28, left: '50%', transform: 'translateX(-50%)',
        zIndex: 100003, background: '#1e293b', color: '#fff',
        borderRadius: 12, boxShadow: '0 8px 32px rgba(0,0,0,0.28)',
        padding: isMobile ? '12px' : '16px 24px', display: 'flex', alignItems: 'center', gap: 12,
        flexDirection: isMobile ? 'column' : 'row', width: isMobile ? 'calc(100vw - 20px)' : 'auto', minWidth: isMobile ? 0 : 360, maxWidth: isMobile ? 'calc(100vw - 20px)' : 520, fontFamily: "'Segoe UI', sans-serif",
        border: '1.5px solid #ef4444',
      }}>
        <div style={{
          width: 36, height: 36, borderRadius: 10, background: '#fee2e2',
          display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0,
        }}>🗑</div>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: '#fca5a5' }}>Vessel deleted in SharePoint</div>
          <div style={{ fontSize: 12, color: '#cbd5e1', marginTop: 3 }}>
            <strong style={{ color: '#fff' }}>{spoVesselDeletedToast.vesselName}</strong> was removed from SharePoint Online and has been moved to the Recycle Bin.
          </div>
        </div>
        <div style={{ display: 'flex', flexDirection: isMobile ? 'row' : 'column', gap: 6, flexShrink: 0, width: isMobile ? '100%' : 'auto' }}>
          <button
            onClick={() => {
              this.setState({ spoVesselDeletedToast: null });
              void this._goToView('recycle');
            }}
            style={{
              background: '#ef4444', color: '#fff', border: 'none', borderRadius: 7,
              minHeight: 44, flex: isMobile ? 1 : undefined, padding: '5px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
            }}
          >View Recycle Bin</button>
          <button
            onClick={() => this.setState({ spoVesselDeletedToast: null })}
            style={{
              background: 'transparent', color: '#94a3b8', border: '1px solid #475569',
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
    const { itemNames, itemType } = spoDocumentDeletedToast;
    const label = itemType === 'folder' ? 'Folder' : 'File';
    const summary = itemNames.length === 1
      ? itemNames[0]
      : `${itemNames.length} ${itemType}s`;
    return (
      <div style={{
        position: 'fixed', bottom: isMobile ? 10 : 28, left: '50%', transform: 'translateX(-50%)',
        zIndex: 100003, background: '#1e293b', color: '#fff', borderRadius: 12,
        boxShadow: '0 8px 32px rgba(0,0,0,0.28)', padding: isMobile ? '12px' : '16px 24px',
        display: 'flex', alignItems: 'center', gap: 12, flexDirection: isMobile ? 'column' : 'row', width: isMobile ? 'calc(100vw - 20px)' : 'auto', minWidth: isMobile ? 0 : 360, maxWidth: isMobile ? 'calc(100vw - 20px)' : 520,
        fontFamily: "'Segoe UI', sans-serif", border: '1.5px solid #ef4444',
      }}>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0 }}>🗑</div>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: '#fca5a5' }}>{label} moved to Recycle Bin</div>
          <div style={{ fontSize: 12, color: '#cbd5e1', marginTop: 3 }}>
            <strong style={{ color: '#fff' }}>{summary}</strong> was deleted from SharePoint Online and moved to the Recycle Bin.
          </div>
        </div>
        <div style={{ display: 'flex', flexDirection: isMobile ? 'row' : 'column', gap: 6, flexShrink: 0, width: isMobile ? '100%' : 'auto' }}>
          <button onClick={() => { this.setState({ spoDocumentDeletedToast: null }); void this._goToView('recycle'); }} style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 7, minHeight: 44, flex: isMobile ? 1 : undefined, padding: '5px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}>View Recycle Bin</button>
          <button onClick={() => this.setState({ spoDocumentDeletedToast: null })} style={{ background: 'transparent', color: '#94a3b8', border: '1px solid #475569', borderRadius: 7, minHeight: 44, flex: isMobile ? 1 : undefined, padding: '5px 12px', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>Dismiss</button>
        </div>
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
        onClose={() => this.setState({ bulkUploadDialog: null })}
      />
    );
  }

  public render(): React.ReactElement {
    const { sessionExpired, view } = this.state;

    if (sessionExpired) {
      return (
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          minHeight: '100vh', background: '#f8fafc', fontFamily: "'Segoe UI', sans-serif",
        }}>
          <div style={{
            background: '#fff', borderRadius: 12, padding: '40px 48px', maxWidth: 420,
            textAlign: 'center', boxShadow: '0 4px 24px rgba(0,0,0,0.10)', border: '1px solid #e2e8f0',
          }}>
            <div style={{ fontSize: 48, marginBottom: 16 }}>🔒</div>
            <h2 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Session Expired</h2>
            <p style={{ margin: '0 0 24px', fontSize: 14, color: '#64748b', lineHeight: 1.6 }}>
              Your session has expired or is no longer valid. Please sign out and sign back in to continue.
            </p>
            <button
              onClick={this._handleSignOut}
              style={{
                background: '#0078d4', color: '#fff', border: 'none', borderRadius: 8,
                padding: '10px 28px', fontSize: 14, fontWeight: 600, cursor: 'pointer',
              }}
            >
              Sign Out & Reload
            </button>
          </div>
        </div>
      );
    }

    let content: React.ReactElement;
    switch (view) {
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
      case 'templates':
        content = this._renderTemplatesPage();
        break;
      case 'approvals':
        content = this._renderApprovalsPage();
        break;
      case 'reports':
        content = this._renderReportsPage();
        break;
      case 'users':
        content = this._renderUsersPage();
        break;
      case 'settings':
        content = this._renderSettingsPage();
        break;
      case 'bento_email':
      case 'email_notify':
        content = this._renderBentoEmailDashboardPage();
        break;
      case 'recycle':
        content = this._renderRecycleBinPage();
        break;
      case 'archive':
        content = this._renderArchivePage();
        break;
      case 'alerts':
        content = renderAlertsPage(this);
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
        {this._renderFolderDeleteDialog()}
        {this._renderSpoVesselDeletedToast()}
        {this._renderSpoDocumentDeletedToast()}
        {renderClassifyDialog(this)}
        {renderVesselSuggestionsModal(this)}
      </>
    );
  }
}
