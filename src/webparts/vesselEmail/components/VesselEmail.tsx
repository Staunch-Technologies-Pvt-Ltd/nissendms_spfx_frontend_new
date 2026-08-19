import * as React from 'react';
import type { IVesselEmailProps } from './IVesselEmailProps';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from './vesselImagePool';
import { createVesselFolders, FolderResult } from './graphFolderService';
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
} from './types/ui';


import {
  cleanName, INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES,
  INITIAL_MOCK_USERS,
} from './constants';
import { renderSidebar } from './pages/Sidebar';
import { renderLayout } from './pages/AppLayout';
import { renderDocPreviewDrawer } from './pages/DocPreviewDrawer';
import { renderDashboard } from './pages/DashboardPage';
import { renderDocumentsPage } from './pages/DocumentsPage';
import { renderVesselsPage, renderFileAlertDialog } from './pages/VesselsPage';
import { renderTemplatesPage } from './pages/TemplatesPage';
import { renderApprovalsPage } from './pages/ApprovalsPage';
import { renderReportsPage } from './pages/ReportsPage';
import { renderUsersPage } from './pages/UsersPage';
import { renderSettingsPage } from './pages/SettingsPage';
import { renderBentoEmailDashboardPage } from './pages/BentoEmailDashboardPage';
import { renderRecycleBinPage } from './pages/RecycleBinPage';
import { renderArchivePage } from './pages/ArchivePage';
import { renderBentoComposeModal } from './modals/BentoComposeModal';
import { renderVesselForm } from './modals/VesselFormModal';
import { renderDeleteModal } from './modals/DeleteVesselModal';

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
  sort: 'default' | 'name_asc' | 'name_desc';
  uploadingGroupKey: string | null;
  uploadInfo: string | null;
  uploadError: string | null;
  selectedFileIds: Set<string>;
  page: number;

  // Modal
  modal: ModalMode;
  selectedVessel: VesselRecord | null;
  deleteVesselIds: Set<string>;
  form: FormState;
  modalBusy: boolean;
  modalMsg: string | null;
  modalError: string | null;
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
  docViewMode: 'folder' | 'list';
  docMainFolder: 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank' | null;
  showAllVesselsInFolderView: boolean;
  docListPage: number;
  docListSort: 'name_az' | 'newest' | 'default';
  docGroupFilter: string;
  documentVesselCount: number;
  documentVesselsLoadingMore: boolean;
  docUploadRowKey: string | null;
  docUploadBusy: boolean;
  docUploadMsg: string | null;
  folderPathStack: { id: string; name: string }[];
  uploadedFilesByFolder: Record<string, { name: string; size: string; date: string; pending?: boolean }[]>;
  selectedDocPreview: DocPreviewItem | null;
  templatesList: TemplateItem[];
  approvalsList: ApprovalItem[];
  approvalTab: 'Pending' | 'Approved' | 'Rejected';
  // Top-header alert bell — new folder/vessel creation alerts (replaces bottom-of-module notifications)
  alertsList: AlertItem[];
  alertFilter: 'all' | 'unread';
  alertOpen: boolean;
  usersList: UserItem[];
  userSearch: string;
  userRoleFilter: string;

  // Settings module state
  settingsTab: 'General' | 'Document Settings' | 'Notification Settings' | 'Permission Settings' | 'Integration' | 'Audit Logs';
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
  } | null;
}


const BLANK_FORM: FormState = { name: '', imo: '', shipyard: '', hull_number: '', vessel_type: '' };

const PAGE_SIZE = 50;

// ── Component Definition ────────────────────────────────────────────────────

export default class VesselEmail extends React.Component<IVesselEmailProps, State> {

  public _abort: AbortController | null = null;
  public _filesLoadedForVessels: Set<string> = new Set();
  public _appUploadedFileIds: Set<string> = new Set();
  public _syncScheduler: SyncScheduler | null = null;
  private _rootFoldersEnsured = false;
  public _isLoadingData = false;
  public _deltaReloadTimer: ReturnType<typeof setTimeout> | null = null;
    public _deltaFileRefreshTimer: ReturnType<typeof setTimeout> | null = null;   // ← add this line
  public _alertRefreshTimer: ReturnType<typeof setInterval> | null = null;
  public _initialDocumentFolderRefreshDone = false;
  public _handleResize = (): void => { this.setState({ windowWidth: window.innerWidth }); };

  public constructor(props: IVesselEmailProps) {
    super(props);
    this.state = {
      rows: [],
      vessels: [],
      loading: false, error: null, reloadKey: 0,
      textFilter: '', vesselFilter: 'all', groupFilter: 'all', catFilter: 'all',
      sort: 'default', uploadingGroupKey: null, uploadInfo: null, uploadError: null,
      selectedFileIds: new Set(), page: 0,
           modal: 'none', selectedVessel: null, deleteVesselIds: new Set(), form: { ...BLANK_FORM },
      modalBusy: false, modalMsg: null, modalError: null, formFieldErrors: {},
      view: 'dashboard',
      recycleBin: [], archiveList: [], panelLoading: false,
      vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all',

      documentsList: INITIAL_MOCK_DOCUMENTS,
      docViewMode: 'folder',
      docMainFolder: null,
      showAllVesselsInFolderView: false,
      docListPage: 0,
      docListSort: 'default',
      docGroupFilter: 'all',
      documentVesselCount: 4,
      documentVesselsLoadingMore: false,
      docUploadRowKey: null,
      docUploadBusy: false,
      docUploadMsg: null,
      folderPathStack: [],
      uploadedFilesByFolder: {},
      selectedDocPreview: null,
      templatesList: INITIAL_MOCK_TEMPLATES,
      approvalsList: [],
      approvalTab: 'Pending',
      alertsList: [],
      alertFilter: 'all',
      alertOpen: false,
      usersList: INITIAL_MOCK_USERS,
      userSearch: '', userRoleFilter: 'all',

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
      sessionExpired: false,
      sessionReady: false,

      folderAnomalies: [],
      normalFolders: [],
      spoClassifyDialog: null,
      spoProvisionDialog: null,
      spoAnomalyDismissConfirm: null,
      spoFileAlertDialog: null,
      spoVesselDeletedToast: null,
      spoDeletedVesselIds: new Set<string>(),
      sidebarCollapsed: false,
      windowWidth: typeof window !== 'undefined' ? window.innerWidth : 1200,
      folderNavHistory: [{ folderPathStack: [], docMainFolder: null }],
      folderNavIndex: 0,
      uploadSuccessPopup: null,
      fileDeleteDialog: null,
      listViewSelectedFiles: new Set<string>(),
      folderViewSelectedFiles: new Set<string>(),
    };
  }

  public _dismissAnomaly = (id: number): void => {
    const base = this._base();
    this._fetchJson(`${base}/api/anomalies/${id}`, {
      method: 'PATCH',
      body: JSON.stringify({ resolved: true }),
    }).catch(() => undefined);
    this.setState(prev => ({
      folderAnomalies: prev.folderAnomalies.filter(a => a.id !== id),
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




  public componentDidMount(): void {
    if (this.props.sessionId) {
      this.setState({ sessionReady: true });
      this._loadData();
    } else if (this.props.sessionInitialized) {
      // Session flow already settled with no session_id (stub/no-DB backend) —
      // do NOT wait for a sessionId that will never arrive. Load data now.
      this._loadData();
    } else {
      // Session not yet available (bypass-login still in flight in the webpart).
      // Retry once after 2 s — by then _ensureSession() will have completed
      // one way or another (with or without a session_id).
      setTimeout(() => {
        if (this.props.sessionId) {
          this.setState({ sessionReady: true });
          this._loadData();
        } else if (this.props.sessionInitialized) {
          this._loadData();
        } else {
          // Last-resort fallback: even if the webpart's onInit somehow never
          // settled (e.g. an unhandled error before the finally{} block), still
          // attempt to load data so the vessel list isn't blank forever.
          this._loadData();
        }
      }, 2000);
    }
    this._startDeltaSync();
    this._loadBentoConfig();
    this._alertRefreshTimer = setInterval(() => this._fetchAlerts(), 30000);
    window.addEventListener('resize', this._handleResize);
    document.addEventListener('click', this._handleOutsideClick);
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
    // Call the vessel-specific flat-tree endpoint immediately — returns in ~50ms from DB only.
    const { vesselFilter } = this.state;
    if (ps.vesselFilter !== vesselFilter && vesselFilter !== 'all') {
      if (!this._filesLoadedForVessels.has(vesselFilter)) {
        this._filesLoadedForVessels.add(vesselFilter);
        this._loadVesselRowsFromApi(vesselFilter).catch(() => undefined);
      }
    }
  }

  public componentWillUnmount(): void {
    this._abort?.abort();
    this._syncScheduler?.stop();
    if (this._deltaReloadTimer) clearTimeout(this._deltaReloadTimer);
    if (this._alertRefreshTimer) clearInterval(this._alertRefreshTimer);
    window.removeEventListener('resize', this._handleResize);
    document.removeEventListener('click', this._handleOutsideClick);
  }

  // ── Delta Sync ────────────────────────────────────────────────────────────

  public _startDeltaSync(): void {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    this._syncScheduler = createSyncScheduler(
      graphClient,
      siteId,
      driveId,
      (result: DeltaSyncResult) => this._applyDeltaResult(result),
      30_000,
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
  public _handleSpoVesselDeletion(deletedId: string, vesselFolderNode: import('./deltaSync').SpoFolderNode): void {
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
    if (!isVesselLevel) return;

    // Match by folder name against known vessels
    const vessel = vessels.find(v => normName(v.name) === normName(vesselFolderNode.name));
    if (!vessel) return;

    const now = new Date().toISOString();
    const recycleBinEntry: DeletedNode = {
      id: vessel.id,
      name: vessel.name,
      kind: 'vessel',
      item_type: 'vessel',
      main_folder: 'Vessels',
      original_path: `Vessels/Specific Vessels/${vessel.name}`,
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
  }

  /**
   * Called when delta sync detects a new folder added at the vessel level
   * (under Vessels/Specific Vessels/ or Vessels/) that is not in the DB.
   * Emits a vessel_unrecognised alert and adds a folderAnomalies entry.
   */
  public _handleSpoNewVesselFolder(node: import('./deltaSync').SpoFolderNode): void {
    const { vessels, folderAnomalies, normalFolders, alertsList } = this.state;
    const normName = (s: string): string => cleanName(s).trim().toLowerCase();

    // Skip structural container folders — these are never vessels
    const nodeNameNorm = normName(node.name);
    if (nodeNameNorm === 'specific vessels' || nodeNameNorm === 'vessels' ||
        nodeNameNorm === 'vessel management') {
      return;
    }
    // Skip if already a known vessel
    if (vessels.some(v => normName(v.name) === nodeNameNorm)) {
      console.log('[VesselDMS] _handleSpoNewVesselFolder: SKIP (known vessel):', node.name);
      return;
    }
    // Skip if already classified as a normal folder
    if ((normalFolders || []).some(f => normName(f.name) === nodeNameNorm)) {
      console.log('[VesselDMS] _handleSpoNewVesselFolder: SKIP (normal folder):', node.name);
      return;
    }
    // Skip if already in anomalies
    if (folderAnomalies.some(a => normName(a.name) === nodeNameNorm)) {
      console.log('[VesselDMS] _handleSpoNewVesselFolder: SKIP (already in anomalies):', node.name);
      return;
    }
    // Skip if already in alerts
    if (alertsList.some(a => a.alert_type === 'vessel_unrecognised' && normName(a.folder_name) === nodeNameNorm)) {
      console.log('[VesselDMS] _handleSpoNewVesselFolder: SKIP (already in alerts):', node.name);
      return;
    }

    console.log('[VesselDMS] _handleSpoNewVesselFolder: EMITTING alert for:', node.name, node.serverRelativePath);

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

    this.setState(prev => ({
      folderAnomalies: [...prev.folderAnomalies, anomaly],
      alertsList: [alert, ...prev.alertsList],
      // Auto-open classify dialog if none is open — navigate to vessels view first
      spoClassifyDialog: prev.spoClassifyDialog ? prev.spoClassifyDialog : {
        anomaly,
        provisioning: false,
        done: false,
        error: null,
      },
      view: prev.spoClassifyDialog ? prev.view : 'vessels',
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

    console.log('[VesselDMS] _scanSpoFolderMapForNewVessels: map size=', spoFolderMap.size, 'known vessels=', knownVesselNames.size);

    for (const [, node] of Array.from(spoFolderMap.entries())) {
      if (!node.isFolder || node.deleted) continue;
      const segs = (node.serverRelativePath || '')
        .split('/').map(s => s.trim().toLowerCase()).filter(Boolean);
      const parentSeg = segs[segs.length - 2] || '';
      if (parentSeg !== 'specific vessels' && parentSeg !== 'vessels') continue;
      // Skip the container folders themselves
      const nameNorm = normName(node.name);
      if (nameNorm === 'specific vessels' || nameNorm === 'vessels' || nameNorm === 'vessel management') continue;

      console.log('[VesselDMS] vessel-level folder found:', node.name, '| path:', node.serverRelativePath, '| known:', knownVesselNames.has(normName(node.name)));
      this._handleSpoNewVesselFolder(node);
    }
  }

  /**
   * Called when delta sync detects a new file added anywhere under the
   * Vessels/ tree (vessel root, any sub-folder, or specific vessels path).
   * Emits a file_outside_structure alert and opens the file-alert dialog.
   */
 public _handleSpoNewFile(node: import('./deltaSync').SpoFolderNode): void {
    const { vessels, alertsList, rows } = this.state;
    const normName = (s: string): string => cleanName(s).trim().toLowerCase();

    if (this._appUploadedFileIds.has(node.id)) {
      this._appUploadedFileIds.delete(node.id);
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
      if (s === 'specific vessels' || s === 'vessels') {
        const candidate = segs[i + 1];
        if (candidate && vessels.some(v => normName(v.name) === normName(candidate))) {
          vesselName = candidate;
        }
        break;
      }
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

    // Build subfolder options from known rows for this vessel (all uploadable leaf folders)
    const subFolderOptions = vesselName
      ? Array.from(
          new Map(
            rows
              .filter(r => normName(r.vesselName) === normName(vesselName!) && r.canUpload)
              .map(r => [r.groupKey, { label: r.subFolderPath, groupKey: r.groupKey, uploadFolderId: r.uploadFolderId, subFolderPath: r.subFolderPath }])
          ).values()
        ).slice(0, 40)
      : [];

    const newDialog: import('./types/ui').SpoFileAlertDialog = {
      fileId: node.id,
      fileName: node.name,
      spoPath: node.serverRelativePath,
      vesselName,
      subFolderOptions: [], // populated dynamically in the dialog via host.state.rows
      moving: false,
      moved: false,
      error: null,
    };

    this.setState(prev => ({
      alertsList: [alert, ...prev.alertsList],
      spoFileAlertDialog: (prev.spoFileAlertDialog && !prev.spoFileAlertDialog.moved)
        ? prev.spoFileAlertDialog
        : newDialog,
    }));

    // Ensure rows are loaded for the detected vessel so the dialog dropdowns are populated
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
        if (node?.isFolder) {
          // Defer vessel deletion handling to after setState (needs current vessels state)
          setTimeout(() => this._handleSpoVesselDeletion(id, node), 0);
        }
        removeNodeFromMap(map, id);
      }

      // Process additions/updates — detect new vessel-level folders
      const missingParentIds = new Set<string>();
      for (const node of result.added) {
        const { missingParentId } = mergeNodeIntoMap(map, node);
        if (missingParentId) missingParentIds.add(missingParentId);

        // Detect new folders added at the vessel level in SPO
        // Skip during baseline scan — vessels state may not be loaded yet,
        // causing false positives for every existing vessel folder.
        if (node.isFolder && !result.isBaseline) {
          const segs = (node.serverRelativePath || '')
            .split('/').map(s => s.trim().toLowerCase()).filter(Boolean);
          const parentSeg = segs[segs.length - 2] || '';
          const isVesselLevel = parentSeg === 'specific vessels' || parentSeg === 'vessels';
          if (isVesselLevel) {
            // Defer to after setState so we have the latest vessels state
            setTimeout(() => this._handleSpoNewVesselFolder(node), 0);
          }
        }
        // Detect files uploaded anywhere under the Vessels/ tree
        // (vessel root, specific vessels, or any sub-folder path)
        if (!node.isFolder && !result.isBaseline) {
          const segs = (node.serverRelativePath || '')
            .split('/').map(s => s.trim().toLowerCase()).filter(Boolean);
          const underVessels = segs.some(s => s === 'vessels' || s === 'specific vessels');
          if (underVessels) {
            setTimeout(() => this._handleSpoNewFile(node), 0);
          }
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
          const newFile = {
            name: fileNode.name,
            size: '—',
            date: 'Today',
            pending: false,
            id: fileNode.id,
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
    return h;
  }

  // Upload headers – DO NOT set Content-Type; the browser must set it with the multipart boundary
  public _uploadHeaders(): Record<string, string> {
    const h: Record<string, string> = {};
    const sid = this.props.sessionId;
    if (sid) { h['Authorization'] = `Bearer ${sid}`; h['X-Session-ID'] = sid; }
    if (this.props.userEmail) { h['X-User-Email'] = this.props.userEmail; }
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
          const fallbackUrl = url.replace('https://nk-dms-dev.sg-nissenkaiun.com', 'http://localhost:8000');
          console.warn(`[VesselDMS] Remote API 502/Network error (${err?.message}) — auto-switched primary base to local backend: ${fallbackUrl}`);
          try {
            const r = await fetch(fallbackUrl, opts);
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
      return 'http://localhost:8000';
    }
    return url || 'https://nk-dms-dev.sg-nissenkaiun.com';
  }

  // ── Alert Bell (top-header) ────────────────────────────────────────────────
  // Fetches the new folder-creation alerts (newly created SharePoint Online
  // folders + newly provisioned vessels) and surfaces them under the header bell.

  public _fetchAlerts = (signal?: AbortSignal): void => {
    this._fetchJson(`${this._base()}/api/alerts`, signal)
      .then((res: any) => {
        if (Array.isArray(res)) {
          // Merge backend alerts with locally-generated ones (vessel_unrecognised, vessel_deleted)
          // so that SPO delta-sync alerts are not wiped out by the periodic backend poll.
          this.setState(prev => {
            const backendIds = new Set((res as AlertItem[]).map(a => a.id));
            const localOnly = prev.alertsList.filter(
              a => !backendIds.has(a.id) &&
                (a.alert_type === 'vessel_unrecognised' || a.alert_type === 'vessel_deleted')
            );
            return { alertsList: [...localOnly, ...(res as AlertItem[])] };
          });
        }
      })
      .catch(() => undefined);
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
    this._fetchJson(`${this._base()}/api/alerts/${id}/read`, { method: 'POST' }).catch(() => undefined);
  };

  public _markAllAlertsRead = (): void => {
    this.setState(prev => ({ alertsList: prev.alertsList.map(a => ({ ...a, read: true })) }));
    this._fetchJson(`${this._base()}/api/alerts/read-all`, { method: 'POST' }).catch(() => undefined);
  };

  public _setAlertFilter = (filter: 'all' | 'unread'): void => {
    this.setState({ alertFilter: filter });
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
        } else if (a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'file') {
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
          read: false,
          created_at: a.detected_at,
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
    this._abort?.abort();
    this._abort = new AbortController();
    const signal = this._abort.signal;
    try {
      const base = this._base();
      if (!base) {
        this.setState({ loading: false, rows: [] });
        return;
      }
      if (this.state.rows.length === 0) {
        this.setState({ loading: true, error: null });
      }

      // Kick off non-critical background tasks
      this._fetchAnomalies(signal);
      this._loadNormalFolders();

      // Step 1: Always fetch vessel list from backend database first
      let vesselList: any = null;
      try { vesselList = await this._fetchJson(`${base}/api/vessels`, signal); } catch { /* ignore */ }
      if (signal.aborted) return;

      let vessels: VesselRecord[] = [];
      if (vesselList && Array.isArray(vesselList) && vesselList.length > 0) {
        const { spoDeletedVesselIds, recycleBin } = this.state;
        const deletedNames = new Set(
          recycleBin
            .filter(r => r.kind === 'vessel' || r.item_type === 'vessel')
            .map(r => r.name.toLowerCase())
        );
        vessels = vesselList
          .map((v: any) => ({ ...v, name: cleanName(v.name), status: v.status || 'Active' }))
          .filter((v: any) => !spoDeletedVesselIds.has(v.id) && !deletedNames.has((v.name || '').toLowerCase()));
      }

      // Step 1b: SPO anomaly scan runs fully in the BACKGROUND — does NOT block vessel/document rendering
      const { graphClient, siteId, driveId } = this.props;
      // Keep the Documents list usable when the configured development API is
      // unavailable (such as an untrusted TLS certificate). In that case,
      // discover vessels from Documents / Vessels directly through Graph.
      if (vessels.length === 0 && graphClient && siteId && driveId) {
        const vesselNodes = await this._getGraphChildren(this.VESSEL_ROOT, signal).catch(() => []);
        if (!signal.aborted) {
          vessels = vesselNodes
            .filter(node => node.isFolder && node.name && !/^pool-/i.test(node.name))
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
              'month end reports', 'service agreements', 'registration', 'drawings and manuals',
              'po & invoice', 'incidents', 'crewing', 'to be classified', 'common for all ships',
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
        this.setState({
          vessels,
          rows: initialRows,
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
          this._refreshDocumentVesselFiles(initialVesselNames);
        });
        return;
      }

      // Step 3: Fallback — Graph API folder-walk (only when flat-tree returns nothing)
      if (graphClient && siteId && driveId && vessels.length > 0) {
        try {
          const graphRows = await this._flattenAllViaGraph(vessels, signal);
          if (!signal.aborted) {
            this.setState({ vessels, rows: this._normalize(graphRows), loading: false });
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
          this.setState({ vessels, rows: this._normalize(rows), loading: false });
        }
      } else {
        this.setState({ vessels, rows: [], loading: false });
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
    // UI breadcrumb and the new Documents hierarchy use the same order:
    // Vessel / Main folder / Category / Sub-category.
    return parts.length >= 2
      ? parts.join('/')
      : fallback.replace(/^\/+/, '');
  }

  /**
   * Immediately fetch rows for a specific vessel from the backend
   * (uses the vessel_name-filtered flat-tree endpoint, hits DB only, ~50ms).
   * Merges the result into `rows` so the list view updates right away.
   */
  public async _loadVesselRowsFromApi(vesselName: string): Promise<void> {
    const base = this._base();
    if (!base) return;
    try {
      const url = `${base}/api/vessels/flat-tree?vessel_name=${encodeURIComponent(vesselName)}`;
      const data = await this._fetchJson(url, new AbortController().signal).catch(() => null);
      if (!data || !Array.isArray(data) || data.length === 0) return;
      const incoming: FlatRow[] = this._normalize(data);
      this.setState(prev => {
        // Replace/merge: remove old rows for this vessel, then prepend fresh ones
        const normV = vesselName.trim().toLowerCase();
        const kept = prev.rows.filter(r => (r.vesselName || '').trim().toLowerCase() !== normV);
        return { rows: [...incoming, ...kept] };
      }, () => {
        // Reset the run-once guard so the new vessel's rows are always refreshed.
        this._lastRefreshedRowsKey = '';
        this._refreshFilesFromBackendRows();
        // DB rows describe the folder structure, but existing files live in
        // SharePoint. Walk the active drive so a refresh never depends on
        // stale cached folder IDs.
        void this._mergeLiveSharePointFiles([vesselName]).catch(err =>
          console.warn('[VesselDMS] live vessel file refresh warning:', err)
        );
      });
    } catch (err) {
      console.warn('[VesselDMS] _loadVesselRowsFromApi warning:', err);
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
        // A vessel can be selected independently while this request is in
        // flight. Replace rows only for this page and retain all others.
        const retained = prev.rows.filter(row =>
          !nextVesselNames.has(cleanName(row.vesselName).trim().toLowerCase())
        );
        return {
          rows: [...retained, ...incoming],
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
  public _refreshDocumentVesselFiles(vesselNames: string[]): void {
    const uniqueVessels = Array.from(new Set(vesselNames.filter(Boolean)));
    if (uniqueVessels.length > 0) void this._mergeLiveSharePointFiles(uniqueVessels);
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

  public _refreshFilesFromBackendRows(): void {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    const { rows, documentVesselCount, vessels } = this.state;
    const visibleVesselNames = new Set(
      vessels.slice(0, documentVesselCount).map(v => cleanName(v.name).trim().toLowerCase())
    );

    // Collect unique subFolderPaths + uploadFolderIds for visible vessels.
    const seen = new Set<string>();
    const toRefresh: Array<{ subFolderPath: string; groupKey: string; uploadFolderId: string }> = [];
    for (const row of rows) {
      if (!visibleVesselNames.has(cleanName(row.vesselName).trim().toLowerCase())) continue;
      if (!row.subFolderPath || seen.has(row.subFolderPath)) continue;
      seen.add(row.subFolderPath);
      toRefresh.push({ subFolderPath: row.subFolderPath, groupKey: row.groupKey, uploadFolderId: row.uploadFolderId });
    }

    if (toRefresh.length === 0) return;

    // Deduplicate runs: skip if the exact same set of folders was already refreshed
    const rowsKey = toRefresh.map(r => r.subFolderPath).sort().join('|');
    if (rowsKey === this._lastRefreshedRowsKey) return;
    this._lastRefreshedRowsKey = rowsKey;

    const BATCH = 5;
    const runBatch = async (items: typeof toRefresh): Promise<void> => {
      for (let i = 0; i < items.length; i += BATCH) {
        // Collect results for the whole batch first, then apply in ONE setState
        // call to prevent concurrent setState calls from overwriting each other.
        const batchResults: Array<{
          groupKey: string;
          uploadFolderId: string;
          parsedFiles: Array<{ name: string; size: string; date: string; pending: boolean; id: string }>;
        }> = [];

        await Promise.all(
          items.slice(i, i + BATCH).map(async item => {
            let fileItems: any[] = [];

            // Strategy 1: Resolve folder via live spoFolderMap (delta-synced active drive).
            // This is the ONLY reliable approach — avoids 404s from stale DB IDs or mismatched paths.
            const liveFolderId = this._getLiveSharePointFolderId(item.subFolderPath);
            if (liveFolderId) {
              const node = this.state.spoFolderMap.get(liveFolderId);
              const memFiles = (node?.children || []).filter(c => !c.isFolder);
              if (memFiles.length > 0) {
                fileItems = memFiles.map(c => ({
                  id: c.id,
                  name: c.name,
                  size: 0,
                  lastModifiedDateTime: null,
                }));
              } else {
                // Node is in the map but children not cached yet — fetch via confirmed live ID
                try {
                  const url = `/sites/${siteId}/drives/${driveId}/items/${liveFolderId}/children?$select=id,name,size,lastModifiedDateTime,file&$top=200`;
                  const result: any = await graphClient!.api(url).get();
                  const hits: any[] = (result?.value ?? []).filter((itm: any) => !!itm.file);
                  if (hits.length > 0) { fileItems = hits; }
                } catch { /* skip */ }
              }
            }

            // Strategy 2: uploadFolderId confirmed in active drive spoFolderMap (fallback).
            // Only safe when the ID is known to be from the current drive — never use raw DB IDs.
            if (fileItems.length === 0 && item.uploadFolderId &&
                this.state.spoFolderMap.has(item.uploadFolderId)) {
              try {
                const url = `/sites/${siteId}/drives/${driveId}/items/${item.uploadFolderId}/children?$select=id,name,size,lastModifiedDateTime,file&$top=200`;
                const result: any = await graphClient!.api(url).get();
                const hits: any[] = (result?.value ?? []).filter((itm: any) => !!itm.file);
                if (hits.length > 0) { fileItems = hits; }
              } catch { /* 404 = folder not in active drive, skip */ }
            }

            // Initial page loads can run before delta sync populates
            // spoFolderMap. Resolve the current folder by its stable path so
            // existing SharePoint files are still shown after a refresh.
            if (fileItems.length === 0) {
              const sharePointFolderPath = this._sharePointFolderPath(item.subFolderPath, '');
              const candidatePath = `${this.VESSEL_ROOT}/Specific Vessels/${sharePointFolderPath}`;
              if (candidatePath) {
                try {
                  const encodedPath = candidatePath.split('/').map(part => encodeURIComponent(part)).join('/');
                  const url = `/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/children?$select=id,name,size,lastModifiedDateTime,file&$top=200`;
                  const result: any = await graphClient!.api(url).get();
                  const hits: any[] = (result?.value ?? []).filter((itm: any) => !!itm.file);
                  if (hits.length > 0) {
                    fileItems = hits;
                  }
                } catch { /* folder may not exist in the active drive */ }
              }
            }

            if (fileItems.length === 0) return;

            batchResults.push({
              groupKey: item.groupKey,
              uploadFolderId: item.uploadFolderId,
              parsedFiles: fileItems.map((f: any) => ({
                name: f.name,
                size: f.size ? `${(f.size / 1024).toFixed(1)} KB` : '—',
                date: f.lastModifiedDateTime ? new Date(f.lastModifiedDateTime).toLocaleString([], { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : 'Today',
                pending: false,
                id: f.id,
              })),
            });
          })
        );

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
    if (!graphClient || !siteId || !driveId || vesselNames.length === 0) return;

    const signal = new AbortController().signal;
    const liveFolderRows: FlatRow[] = [];
    await this._mapLimit(vesselNames.filter(Boolean), 2, async vesselName => {
      const rows = await this._flattenVesselViaGraph(vesselName, signal, () => undefined).catch(() => []);
      // _walkGraphFolder emits one row for every uploadable leaf, even when
      // it has no files. Keeping those empty rows is essential: their
      // uploadFolderId is the actual SPO location used for later uploads.
      liveFolderRows.push(...rows.filter(row => Boolean(row.uploadFolderId)));
    });
    if (liveFolderRows.length === 0) return;

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
            size: '—',
            date: row.fileUploadedAt ? new Date(row.fileUploadedAt).toLocaleString() : 'Today',
            pending: false,
            id: row.fileId || row.fileName,
            uploadedAt: row.fileUploadedAt,
          };
          if (row.groupKey) {
            const list = updatedByFolder[row.groupKey] || [];
            if (!list.some((f: any) => f.name === row.fileName)) {
              updatedByFolder[row.groupKey] = [...list, newF];
            }
          }
          if (row.uploadFolderId) {
            const list = updatedByFolder[row.uploadFolderId] || [];
            if (!list.some((f: any) => f.name === row.fileName)) {
              updatedByFolder[row.uploadFolderId] = [...list, newF];
            }
          }
        }
      });

      return { rows: [...retained, ...mergedFolderRows], uploadedFilesByFolder: updatedByFolder };
    });
  }

  /** Open the exact SharePoint folder represented by a Documents list row. */
  public async _openSharePointFolder(row: GroupedRow): Promise<void> {
    // Open synchronously so the browser does not block the user-initiated tab
    // while the Graph request resolves the folder's SharePoint web URL.
    const target = window.open('', '_blank', 'noopener,noreferrer');
    const { graphClient, siteId, driveId } = this.props;
    const normalisePath = (path: string): string =>
      (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
    const sharePointPath = this._sharePointFolderPath(row.subFolderPath, '');
    const expectedPath = normalisePath(sharePointPath);
    let liveFolderId: string | undefined;

    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder) continue;
      const path = normalisePath(node.serverRelativePath);
      if (path === expectedPath || (expectedPath && path.endsWith(`/${expectedPath}`))) {
        liveFolderId = node.id;
        break;
      }
    }

    try {
      if (!graphClient || !siteId || !driveId) {
        throw new Error('SharePoint connection is not available for this web part.');
      }

      let item: any;
      if (liveFolderId) {
        item = await graphClient
          .api(`/sites/${siteId}/drives/${driveId}/items/${liveFolderId}?$select=id,folder,webUrl`)
          .get();
      } else {
        // The delta baseline may still be loading on the first click. Resolve
        // the breadcrumb directly in Graph instead of asking the user to
        // refresh. Some libraries keep vessel folders below this container.
        const paths = [
          `Vessels/Specific Vessels/${sharePointPath}`,
          `Vessels/${sharePointPath}`,
          sharePointPath,
          sharePointPath.replace(/\/Vessels\//i, '/'),
        ];
        if (sharePointPath && !sharePointPath.toLocaleLowerCase().startsWith('vessel management/')) {
          paths.push(`Vessel Management/${sharePointPath}`);
          paths.push(`Vessel Management/Vessels/Specific Vessels/${sharePointPath}`);
        }
        let lastError: any;
        for (const candidatePath of paths.filter(Boolean)) {
          try {
            const encodedPath = candidatePath.split('/').map(part => encodeURIComponent(part)).join('/');
            const candidate: any = await graphClient
              .api(`/sites/${siteId}/drives/${driveId}/root:/${encodedPath}?$select=id,folder,webUrl`)
              .get();
            if (candidate?.folder && candidate?.webUrl) {
              item = candidate;
              break;
            }
          } catch (pathError) {
            lastError = pathError;
          }
        }
        if (!item) {
          throw lastError || new Error('The selected SharePoint folder was not found.');
        }
      }
      if (!item?.webUrl) throw new Error('SharePoint did not return a folder link.');
      if (target) target.location.href = item.webUrl;
      else window.open(item.webUrl, '_blank', 'noopener,noreferrer');
    } catch (err: any) {
      if (target) target.close();
      console.warn('[VesselDMS] failed to open SharePoint folder:', err);
      alert(err?.message || 'Unable to open this SharePoint folder.');
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

    const cleanKey = (s: string) => (s || '').toLowerCase()
      .replace(/^folder-\d+\s+/i, '')
      .replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '')
      .replace(/[^\w\d]/g, '');

    const cleanLeaf = cleanKey(parts[parts.length - 1]);
    const cleanVessel = cleanKey(parts[0]);
    const cleanGroup = parts.length > 1 ? cleanKey(parts[1]) : '';
    const cleanParent = parts.length >= 2 ? cleanKey(parts[parts.length - 2]) : '';

    const legacyPath = parts.length >= 3
      ? normalisePath([parts[1], parts[0], ...parts.slice(2)].join('/'))
      : '';

    // First pass: exact or suffix path matches (fastest and most specific)
    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder) continue;
      const livePath = normalisePath(node.serverRelativePath);
      if (expectedPath && (livePath === expectedPath || livePath.endsWith(`/${expectedPath}`))) return node.id;
      if (legacyPath && (livePath === legacyPath || livePath.endsWith(`/${legacyPath}`))) return node.id;
    }

    // Second pass: segment token matching (handles custom folder prefixes and structure variations)
    // Require the live path to have at least as many segments as the breadcrumb to avoid
    // matching a shallow parent folder (e.g. "Technical & Crewing") instead of a leaf.
    const minSegments = parts.length;
    for (const [, node] of Array.from(this.state.spoFolderMap.entries())) {
      if (!node.isFolder) continue;
      const nodeSegments = (node.serverRelativePath || '').split('/').filter(Boolean).map(cleanKey);
      if (nodeSegments.length < minSegments) continue;
      const lastSeg = nodeSegments[nodeSegments.length - 1];

      if (lastSeg === cleanLeaf || (lastSeg && cleanLeaf && (lastSeg.includes(cleanLeaf) || cleanLeaf.includes(lastSeg)))) {
          // Require the live folder's immediate parent segment to match the
        // breadcrumb's immediate parent, so "To be Classified" directly under
        // "Commercial & Chartering" is never confused with the nested
        // "Commercial & Chartering > Agreements > To be Classified".
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
  ): Promise<{ fileId: string | null; statusPending: boolean; folderId: string | null; isGraphUpload: boolean }> {
    const base = this._base();
    if (!base) return { fileId: null, statusPending: false, folderId: null, isGraphUpload: false };

    // A backend folder ID may belong to its configured drive rather than the
    // Documents library hosting this web part. For regular uploads, build the
    // visible SPO destination from the breadcrumb instead:
    // Vessel > Main folder > Category > Sub-category.
    const sharePointFolderPath = this._sharePointFolderPath(subFolderPath, uploadFolderId);
    const liveFolderId = this._getLiveSharePointFolderId(subFolderPath);
    // A live list row carries the exact path that was discovered from SPO. It
    // may be in the legacy Main-folder/Vessel layout, which differs from the
    // UI breadcrumb's new Vessels/Vessel layout.
    const resolvedSharePointFolderPath = uploadFolderId.includes('/')
      ? uploadFolderId.replace(/^\/+|\/+$/g, '')
      : sharePointFolderPath;

    // Write documents directly to SharePoint via Graph API when available.
    // This avoids the backend OCR requirement (month-upload endpoint) and ensures
    // the file lands in the exact visible SPO folder shown in the list.
    const { graphClient, siteId, driveId } = this.props;
    if (graphClient && siteId && driveId) {
      try {
        let folder: any = null;

        // 1. If we have a confirmed live folder ID in the active drive, use it directly
        if (liveFolderId) {
          try {
            folder = await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${liveFolderId}?$select=id,name,folder,webUrl`).get();
          } catch {
            folder = null;
          }
        }

        // 2. If uploadFolderId is an item ID in spoFolderMap, try it
        if (!folder && uploadFolderId && !uploadFolderId.includes('/') && this.state.spoFolderMap.has(uploadFolderId)) {
          try {
            folder = await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${uploadFolderId}?$select=id,name,folder,webUrl`).get();
          } catch {
            folder = null;
          }
        }

        // 3. Candidate path resolution (checks modern specific vessels, legacy, and container layouts)
        if (!folder) {
          const parts = (subFolderPath || '').split('>').map(p => p.trim()).filter(Boolean);
          const vName = vesselName || (parts.length > 0 ? parts[0] : '');
          const grp = parts.length > 1 ? parts[1] : '';
          const rest = parts.length > 2 ? parts.slice(2).join('/') : '';
          const cleanSharePointPath = sharePointFolderPath.replace(/^\/+|\/+$/g, '');

          const candidatePaths = [
            cleanSharePointPath ? `${this.VESSEL_ROOT}/Specific Vessels/${cleanSharePointPath}` : '',
            cleanSharePointPath ? `${this.VESSEL_ROOT}/${cleanSharePointPath}` : '',
            vName && grp ? `${this.VESSEL_ROOT}/Specific Vessels/${vName}/${grp}${rest ? `/${rest}` : ''}` : '',
            vName && grp ? `${this.VESSEL_ROOT}/${vName}/${grp}${rest ? `/${rest}` : ''}` : '',
            grp && vName ? `${grp}/${vName}${rest ? `/${rest}` : ''}` : '',
            cleanSharePointPath,
            resolvedSharePointFolderPath,
            cleanSharePointPath ? `Vessel Management/${cleanSharePointPath}` : '',
            cleanSharePointPath ? `Vessel Management/Vessels/Specific Vessels/${cleanSharePointPath}` : '',
          ].filter(Boolean);

          for (const tryPath of candidatePaths) {
            try {
              const encodedPath = tryPath.split('/').map(s => encodeURIComponent(s)).join('/');
              const candidate = await graphClient.api(`/sites/${siteId}/drives/${driveId}/root:/${encodedPath}?$select=id,name,folder,webUrl`).get();
              if (candidate?.id && candidate?.folder) {
                folder = candidate;
                break;
              }
            } catch {
              // try next candidate
            }
          }
        }

        if (!folder?.id || !folder?.folder) {
          throw new Error('The selected SharePoint folder was not found.');
        }

       const uploadUrl = `/sites/${siteId}/drives/${driveId}/items/${folder.id}:/${encodeURIComponent(file.name)}:/content`;
        const item: any = await graphClient.api(uploadUrl).put(file);
        if (!item?.id) throw new Error('SharePoint did not return an uploaded file ID.');
        this._appUploadedFileIds.add(item.id as string);
        // Graph upload: folderId is a real SharePoint drive item ID (safe for Graph refresh)
        return { fileId: item.id as string, statusPending: false, folderId: folder.id as string, isGraphUpload: true };
      } catch (graphErr) {
        const detail = graphErr instanceof Error ? graphErr.message : String(graphErr);
        throw new Error(`Could not upload to SharePoint folder: ${detail}`);
      }
    }

    // REST backend fallback (no Graph client configured).
    const form = new FormData();
    form.append('file', file);
    if (this.props.userEmail) form.append('uploader_email', this.props.userEmail);

    // Prefer the live confirmed drive folder ID for the REST endpoint to avoid
    // stale DB drive_item_ids that belong to a different drive.
    const restFolderId = liveFolderId || uploadFolderId;

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

  // ── Graph API Folder Walking (mirrors VesselListView.tsx from reference project) ────

  public readonly GRAPH_FETCH_CONCURRENCY = 6;
  // Current SharePoint structure: Documents / Vessels / {Vessel} / {Main Folder}.
  // Legacy layouts are still checked below as fallbacks for existing content.
  public readonly VESSEL_ROOT = 'Vessels';
  public readonly MAIN_FOLDER_NAMES = [
    'Technical & Crewing',
    'Commercial & Chartering',
    'Insurance',
    'Kaizen - Knowledge Bank',
  ];

  /**
   * List children of a SharePoint drive path via Graph API.
   * Returns an array of { id, name, isFolder, monthDriven, upload } items.
   */
  public async _getGraphChildren(
    folderPath: string,
    signal: AbortSignal,
  ): Promise<Array<{ id: string; name: string; isFolder: boolean; monthDriven: boolean; upload: boolean; lastModifiedDateTime?: string }>> {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return [];

    const fetchChildrenForPath = async (p: string) => {
      const cleanP = p.replace(/^\/+/, '');
      const encodedPath = cleanP.split('/').map(s => encodeURIComponent(s)).join('/');
      // Graph root children uses /root/children, not /root:/:/children.
      const itemPath = cleanP ? `root:/${encodedPath}:/children` : 'root/children';
      const url = `/sites/${siteId}/drives/${driveId}/${itemPath}` +
        `?$select=id,name,folder,file,lastModifiedDateTime&$top=200`;
      const result: any = await graphClient.api(url).get();
      if (signal.aborted) return [];
      return (result.value ?? []).map((item: any) => ({
        id: item.id as string,
        name: item.name as string,
        isFolder: !!item.folder,
        monthDriven: false,
        upload: !!item.folder,
        lastModifiedDateTime: item.lastModifiedDateTime,
      }));
    };

    // Only try the Vessel Management/ prefix for shallow paths (≤2 segments).
    // Deep sub-folder paths that 404 simply don't exist in that layout and
    // probing them doubles every Graph request during a vessel walk.
    const segmentCount = folderPath ? folderPath.split('/').length : 0;
    const tryVmPrefix = !folderPath.startsWith('Vessel Management') && segmentCount <= 2;

    try {
      const kids = await fetchChildrenForPath(folderPath);
      if (kids.length > 0) return kids;
      if (tryVmPrefix) {
        const vmKids = await fetchChildrenForPath(`Vessel Management/${folderPath}`).catch(() => []);
        if (vmKids.length > 0) return vmKids;
      }
      return kids;
    } catch (err: any) {
      const status = err?.statusCode ?? err?.code ?? 0;
      if (status === 404 && tryVmPrefix) {
        try {
          return await fetchChildrenForPath(`Vessel Management/${folderPath}`);
        } catch {
          return [];
        }
      }
      if (status === 404) return [];
      throw err;
    }
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
  ): Promise<void> {
    if (signal.aborted) return;

    let kids: Array<{ id: string; name: string; isFolder: boolean; monthDriven: boolean; upload: boolean; lastModifiedDateTime?: string }>;
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

    // DMS folder structure:
    //   pathParts[0] = Main Folder (Group):  "Technical & Crewing", "Commercial & Chartering", etc.
    //   pathParts[1] = Vessel Name:           "MV Pacific Test 2"
    //   pathParts[2] = Category:              "Registration", "Month End Reports", etc.
    //   pathParts[3] = Sub-Category (leaf):   "Flag & MPA", "Main Engine", etc.
    const group = pathParts.length >= 1 ? stripPrefix(pathParts[0]) : '';
    // Category = folder immediately under vessel (index 2)
    const category = pathParts.length >= 3 ? stripPrefix(pathParts[2]) : (pathParts.length >= 2 ? stripPrefix(pathParts[pathParts.length - 1]) : group);
    // Sub-category = deepest leaf (index 3+). If no deeper level, equals category
    const subCategory = pathParts.length >= 4 ? stripPrefix(pathParts[pathParts.length - 1]) : category;

    // Breadcrumb: VesselName > Group > Category > SubCategory
    const breadcrumbParts = [vesselName, group, category];
    if (subCategory && subCategory !== category) breadcrumbParts.push(subCategory);
    const subPath = breadcrumbParts.join(' > ');
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
          uploadFolderId: folderPath,
          monthDriven: false,
        }]);
      } else {
        const leafRows: FlatRow[] = files.map((f, idx) => ({
          srNo: idx === 0 ? baseSr : `${baseSr}${suffixes[idx - 1] ?? idx}`,
          vesselName, group, category, subCategory, subFolderPath: subPath,
          fileName: f.name, fileId: f.id, canUpload, groupKey,
          uploadFolderId: folderPath,
          fileUploadedAt: f.lastModifiedDateTime ? Date.parse(f.lastModifiedDateTime) : undefined,
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
    const basePaths = ['', 'Vessel Management'];

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
            discoveredMainFolders.push({ path: `${vesselPath}/${node.name}`, group: node.name });
          });
        }
      }

      // Legacy structure: {Main folder} / {Vessel} / ... or
      // {Main folder} / Vessels / {Vessel} / ....
      if (discoveredMainFolders.length === 0) {
        const mainNodes = baseChildren.filter(node => node.isFolder && knownMainFolders.has(normaliseName(node.name)));
        if (mainNodes.length > 0) foundSupportedStructure = true;
        for (const mainNode of mainNodes) {
          if (signal.aborted) break;
          const mainPath = `${prefix}${mainNode.name}`;
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
          if (vesselNode && vesselPath) discoveredMainFolders.push({ path: vesselPath, group: mainNode.name });
        }
      }
    }

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
          const templateTree: Record<string, Array<{ category: string; subCats: string[] }>> = {
            'Technical & Crewing': [
              { category: 'Month End Reports', subCats: ['Main Engine', 'Aux Engine', 'Cooling Water', 'Inspection Reports', 'Defect Reports', 'Guarantee Claims', 'To be Classified'] },
              { category: 'Service Agreements', subCats: ['Technical Management', 'Crew Management', 'Vendor & Service Provider', 'To be Classified'] },
              { category: 'Registration', subCats: ['Flag & MPA', 'Ship Builder', 'Radio & Telecom', 'Crewing & SMOU', 'Novation', 'To be Classified'] },
              { category: 'Drawings and Manuals', subCats: ['Drawing', 'Manual', 'To be Classified'] },
              { category: 'PO & Invoice', subCats: ['Purchase Order', 'Vendor Invoice'] },              { category: 'Incidents', subCats: ['Incidents'] },
              { category: 'Crewing', subCats: ['Crewing'] },
              { category: 'To be Classified', subCats: ['To be Classified'] },
            ],
            'Commercial & Chartering': [
              { category: 'Agreements', subCats: ['Charter party', 'Pool Agreement', 'Commission Agreement', 'To be Classified'] },
              { category: 'Invoices & Payments', subCats: ['Invoice', 'Payments', 'To be Classified'] },
              { category: 'Claims & Disputes', subCats: ['Disputes', 'Claims', 'To be Classified'] },
              { category: 'To be Classified', subCats: ['To be Classified'] },
            ],
            'Insurance': [
              { category: 'P&I', subCats: ['P&I'] },
              { category: 'H&M', subCats: ['H&M'] },
              { category: 'War Risk', subCats: ['War Risk'] },
              { category: 'Flag & MPA', subCats: ['Flag & MPA'] },
              { category: 'USA Related', subCats: ['USA Related'] },
            ],
            'Kaizen - Knowledge Bank': [
              { category: 'Circulars and Guidance', subCats: ['Equipment Maker', 'Class', 'Flag / Port State', 'SIRE/OCIMF/RightShip', 'Shipyard'] },
              { category: 'Lessons Learned', subCats: ['Lessons Learned'] },
              { category: 'Procedures and Work Instructions', subCats: ['Procedures and Work Instructions'] },
              { category: 'Templates', subCats: ['Templates'] },
            ],
          };

          const cats = templateTree[displayMainFolder] || [
            { category: displayMainFolder, subCats: [displayMainFolder] }
          ];

          const fallbackRows: FlatRow[] = [];
          for (const item of cats) {
            for (const subCat of item.subCats) {
              srCounter.value += 1;
              const breadcrumb = subCat !== item.category
                ? `${vesselName} > ${displayMainFolder} > ${item.category} > ${subCat}`
                : `${vesselName} > ${displayMainFolder} > ${item.category}`;
              const logicalPath = `${vesselName}/${displayMainFolder}/${item.category}/${subCat}`;
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
    this.setState({ view });
    if (view === 'vessels') {
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
      this.setState({ panelLoading: true });
      try {
        const data = await this._fetchJson(`${this._base()}/api/recycle-bin/nodes`);
        const backendItems = (data || []).map((v: any) => ({ ...v, name: cleanName(v.name || '') }));
        this.setState(prev => {
          // Preserve locally-deleted files (deleted via Graph, not tracked by backend)
          const localFiles = prev.recycleBin.filter(item => item.kind === 'file' || item.item_type === 'file');
          const backendIds = new Set(backendItems.map((item: any) => item.id));
          const newLocalFiles = localFiles.filter(item => !backendIds.has(item.id));
          return { recycleBin: [...backendItems, ...newLocalFiles], panelLoading: false };
        });
      } catch {
        this.setState(prev => {
          const localFiles = prev.recycleBin.filter(item => item.kind === 'file' || item.item_type === 'file');
          return { recycleBin: localFiles, panelLoading: false };
        });
      }
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

        const mapped: ApprovalItem[] = combined.map((a: any) => ({
          id: String(a.id || a._id || Date.now()),
          documentName: a.filename || a.file_name || a.fileName || a.document_name || a.documentName || a.name || 'Unknown',
          vessel: a.vessel_name || a.vesselName || a.vessel || '—',
          requestedBy: a.uploaded_by_email || a.uploaded_by_name || a.requested_by || a.requestedBy || a.uploader_email || a.uploader || a.user_email || '—',
          requestedOn: mapDate(a),
          status: a.status === 'approved' ? 'Approved' : a.status === 'rejected' ? 'Rejected' : 'Pending',
        }));
        // Merge: keep any locally-added pending items not yet in the backend response
        const backendIds = new Set(mapped.map(m => m.id));
        const localOnly = this.state.approvalsList.filter(a => !backendIds.has(a.id));
        this.setState({ approvalsList: [...mapped, ...localOnly], panelLoading: false });
      } catch {
        this.setState({ panelLoading: false });
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
      selectedVessel: v,
      form: {
        name: v.name, imo: v.imo || '',
        shipyard: v.shipyard || '', hull_number: v.hull_number || '',
        vessel_type: v.vessel_type || '',
      },
      modalMsg: null, modalError: null,
    });
  };

  public _openDeleteVessel = (v?: VesselRecord): void => {
    const selected = v || this.state.selectedVessel;
    this.setState({
      modal: 'delete',
      selectedVessel: selected || null,
      deleteVesselIds: selected ? new Set([selected.id]) : new Set(),
      modalMsg: null,
      modalError: null,
    });
  };

  public _closeModal = (): void => {
    if (!this.state.modalBusy) this.setState({ modal: 'none', modalMsg: null, modalError: null });
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
    };

        try {
      const res = await fetch(`${this._base()}/api/vessels`, {
        method: 'POST', headers: this._headers(),
        body: JSON.stringify({ name: form.name.trim(), imo: form.imo.trim(), shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null }),
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
      }));

      // Fire-and-forget folder provisioning — backend already handles SPO folder creation in background
      void this._provisionVesselFolders(form.name.trim(), newVesselRecord.id).catch(() => undefined);

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
      }));
    }
  };

  public _submitEdit = async (): Promise<void> => {
    const { form, selectedVessel, vessels } = this.state;
    if (!selectedVessel) return;
    this.setState({ modalBusy: true, modalError: null, modalMsg: null });
    try {
      const res = await fetch(`${this._base()}/api/vessels/${selectedVessel.id}`, {
        method: 'PATCH', headers: this._headers(),
        body: JSON.stringify({ name: form.name.trim() || null, imo: form.imo.trim() || null, shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null }),
      });
      const data = await res.json();
      if (!res.ok && res.status !== 202) throw new Error(data?.message ?? `Error ${res.status}`);
      const msg = data.status === 'pending' ? '⏳ Vessel update submitted for approval.' : `✅ Vessel updated successfully.`;
      this.setState({ modalBusy: false, modalMsg: msg });
      setTimeout(() => this.setState({ modal: 'none', reloadKey: this.state.reloadKey + 1 }), 1600);
    } catch (e: any) {
      const updated = vessels.map(v => v.id === selectedVessel.id ? { ...v, name: form.name.trim() || v.name, imo: form.imo.trim() || v.imo, shipyard: form.shipyard || v.shipyard, hull_number: form.hull_number || v.hull_number, vessel_type: form.vessel_type || v.vessel_type } : v);
      this.setState({ vessels: updated, modalBusy: false, modalMsg: '✅ Vessel updated successfully.' });
      setTimeout(() => this.setState({ modal: 'none' }), 1200);
    }
  };

  public _submitDelete = async (): Promise<void> => {
    const { deleteVesselIds, vessels } = this.state;
    const selected = vessels.filter(v => deleteVesselIds.has(v.id));
    if (selected.length === 0) {
      this.setState({ modalError: 'Select at least one vessel to delete.' });
      return;
    }
    this.setState({ modalBusy: true, modalError: null });
    const deletedIds: string[] = [];
    const pendingNames: string[] = [];
    const failures: string[] = [];

    for (const vessel of selected) {
      try {
        const res = await fetch(`${this._base()}/api/vessels/${vessel.id}?vessel_name=${encodeURIComponent(vessel.name)}`, {
          method: 'DELETE', headers: this._headers(),
        });
        const raw = await res.text();
        let data: any = {};
        try { data = raw ? JSON.parse(raw) : {}; } catch { /* response is not JSON */ }
        if (!res.ok && res.status !== 202) throw new Error(data?.message || raw || `Error ${res.status}`);
        if (data.status === 'pending' || res.status === 202) {
          pendingNames.push(vessel.name);
        } else {
          deletedIds.push(vessel.id);
        }
      } catch (e: any) {
        failures.push(`${vessel.name}: ${e?.message || 'Delete failed'}`);
      }
    }

    const removedCount = deletedIds.length;
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
      selectedVessel: deletedIds.includes(prev.selectedVessel?.id || '') ? null : prev.selectedVessel,
      deleteVesselIds: new Set(pendingNames.length || failures.length ? selected.filter(v => !deletedIds.includes(v.id)).map(v => v.id) : []),
      // Add deleted vessels to recycle bin immediately (deduplicated by name)
      recycleBin: [
        ...immediateRecycleBinEntries,
        ...prev.recycleBin.filter(r => !immediateRecycleBinEntries.some(e => e.name.toLowerCase() === r.name.toLowerCase())),
      ],
    }));

    // Emit an alert for each deleted vessel so it appears in the header bell.
    if (removedCount > 0) {
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
            this.setState({ recycleBin: nodes });
          }
        } catch { /* The item will be visible after the next Recycle Bin refresh. */ }
      };
      void fetchRecycleBin(1);
    }
    if (!failures.length && !pendingNames.length) {
      setTimeout(() => this.setState({ modal: 'none', selectedVessel: null, deleteVesselIds: new Set() }), 1600);
    }
  };

  // SharePoint Folder Provisioning ─────────────────────────────────────────

  public async _provisionVesselFolders(
    vesselName: string,
    vesselId?: string,
  ): Promise<{ success: boolean; results: FolderResult[] }> {
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
      const provisionedSet = new Set(this.state.provisionedVesselIds);
      if (vesselId) {
        provisionedSet.add(vesselId);
        // Persist provisioned status in database
        void fetch(`${this._base()}/api/vessels/${vesselId}/provision`, {
          method: 'POST',
          headers: this._headers(),
        }).catch(() => undefined);
      }
      this.setState({
        folderCreationBusy: false,
        folderProvisioningVesselId: null,
        folderCreationResults: result.results,
        folderCreationError: result.success ? null : 'Some folders could not be created. Check the creation log.',
        provisionedVesselIds: provisionedSet,
      });
      // Trigger an immediate delta sync so the new folders appear in the tree
      void this._syncScheduler?.triggerNow().catch(() => undefined);
      return { success: result.success, results: result.results };
    } catch (err: any) {
      const msg = err?.message ?? 'Folder provisioning failed.';
      this.setState({
        folderCreationBusy: false,
        folderProvisioningVesselId: null,
        folderCreationError: msg,
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
      const BATCH = 5;
      for (let i = 0; i < liveFolders.length; i += BATCH) {
        const batch = liveFolders.slice(i, i + BATCH);
        await Promise.all(batch.map(folder => this._refreshFolderFiles(folder.id, folder.groupKey, true)));
      }
    }

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

      const expectedCats: Record<string, string[]> = {
        'Technical & Crewing': ['Month End Reports', 'Service Agreements', 'Registration', 'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified'],
        'Commercial & Chartering': ['Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified'],
        'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag and MPA'],
        'Kaizen - Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
        'Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
      };
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

  public async _refreshFolderFiles(
    uploadFolderId: string,
    groupKey: string,
    force: boolean = false,
    isConfirmedGraphFolderId: boolean = false,
  ): Promise<void> {
    if (!uploadFolderId) return;
    if (/^f\d+$/.test(uploadFolderId)) return;
    if (!force && this._filesLoadedForFolders.has(uploadFolderId)) return;
    // Prevent concurrent duplicate fetches for the same folder
    if (this._refreshFolderFilesInFlight.has(uploadFolderId)) return;
    this._refreshFolderFilesInFlight.add(uploadFolderId);

    // For path-based IDs (containing '/'), resolve to a real folder ID first
    let resolvedId = uploadFolderId;
    if (uploadFolderId.includes('/')) {
      try {
        const resolved = await this._fetchJson(
          `${this._base()}/api/folders/upload-by-path?path=${encodeURIComponent(uploadFolderId)}&resolve_only=true`
        ).catch(() => null);
        if (resolved?.folder_id) {
          resolvedId = resolved.folder_id;
        } else {
          this._refreshFolderFilesInFlight.delete(uploadFolderId);
          return;
        }
      } catch {
        this._refreshFolderFilesInFlight.delete(uploadFolderId);
        return;
      }
    }

    // Prefer live Graph API fetch so files uploaded directly in SPO are visible immediately
    const { graphClient, siteId, driveId } = this.props;
    let fileItems: any[] = [];
    let fetchedFromGraph = false;

    // First, try to find the matching row using multiple strategies
    // Strategy 1: exact uploadFolderId match
    // Strategy 2: resolvedId match (after path resolution)
    // Strategy 3: groupKey match (for newly uploaded files)
    // Strategy 4: match by subFolderPath breadcrumb if we have a groupKey
    let matchingRow = this.state.rows.find(r =>
      r.uploadFolderId === uploadFolderId ||
      r.uploadFolderId === resolvedId ||
      (groupKey && r.groupKey === groupKey)
    );

    // Strategy 4: if still no match and we have groupKey, try to find by breadcrumb
    if (!matchingRow && groupKey) {
      matchingRow = this.state.rows.find(r => r.groupKey === groupKey);
    }

    const sharePointFolderPath = matchingRow
      ? this._sharePointFolderPath(matchingRow.subFolderPath, '')
      : '';
    // This ID was obtained from the active drive's delta response, unlike a
    // database folder ID which can point to another drive.
    const isLiveDriveFolderId = this.state.spoFolderMap.has(uploadFolderId);

    // Also check if resolvedId is in the live drive folder map
    const isResolvedIdLive = this.state.spoFolderMap.has(resolvedId);

    if (graphClient && siteId && driveId) {
      // Prefer the known Documents-library path. Backend IDs in this project
      // can belong to another drive and produce Graph itemNotFound errors here.
      // Only attempt path lookup when path has ≥3 segments (VesselName/MainFolder/Category).
      const pathSegments = sharePointFolderPath ? sharePointFolderPath.split('/').filter(Boolean) : [];
      if (sharePointFolderPath && !isLiveDriveFolderId && pathSegments.length >= 3) {
        const tryPaths = [
          `${this.VESSEL_ROOT}/Specific Vessels/${sharePointFolderPath}`,
          `${this.VESSEL_ROOT}/${sharePointFolderPath}`,
        ];
        for (const tryPath of tryPaths) {
          try {
            const encodedPath = tryPath.split('/').map(s => encodeURIComponent(s)).join('/');
            const pathUrl = `/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/children?$select=id,name,size,lastModifiedDateTime,file,folder&$top=200`;
            const result: any = await graphClient.api(pathUrl).get();
            const items: any[] = result?.value ?? [];
            const hits = items.filter((i: any) => !i.folder);
            if (hits.length > 0 || (result?.value && Array.isArray(result.value))) {
              fileItems = hits;
              fetchedFromGraph = true;
              console.log(`[VesselDMS] _refreshFolderFiles: fetched ${fileItems.length} files via path lookup for "${tryPath}"`);
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
        targetGraphId = liveIdFromBreadcrumb;
      }
      if (!fetchedFromGraph && targetGraphId && !targetGraphId.includes('/')) {
        try {
          const url = `/sites/${siteId}/drives/${driveId}/items/${targetGraphId}/children?$select=id,name,size,lastModifiedDateTime,file,folder&$top=200`;
          const result: any = await graphClient.api(url).get();
          const items: any[] = result?.value ?? [];
          fileItems = items.filter((i: any) => !i.folder); // only files, not subfolders
          fetchedFromGraph = true;
          console.log(`[VesselDMS] _refreshFolderFiles: fetched ${fileItems.length} files via ID lookup for "${targetGraphId}"`);
        } catch (idLookupError) {
          console.warn(`[VesselDMS] failed to load confirmed live folder ID "${targetGraphId}"`, idLookupError);
        }
      }
    }

    if (!fetchedFromGraph) {
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

    // Keep empty responses retryable because SharePoint can briefly return an
    // empty children collection immediately after an upload.
    if (fileItems.length > 0) this._filesLoadedForFolders.add(uploadFolderId);
    this._refreshFolderFilesInFlight.delete(uploadFolderId);

    const parsedUploads = fileItems.map((f: any) => ({
      name: f.name || f.displayName,
      size: f.size ? `${(f.size / 1024).toFixed(1)} KB` : '—',
      date: f.lastModifiedDateTime ? new Date(f.lastModifiedDateTime).toLocaleString([], { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : (f.modified || 'Today'),
      pending: false,
      id: f.id,
  uploadedAt: f.lastModifiedDateTime ? Date.parse(f.lastModifiedDateTime) : undefined,
    }));

    this.setState(prev => {
      // Merge fetched parsedUploads with any existing local uploads so we don't wipe out freshly uploaded files.
      // Also check groupKey so a just-uploaded file stored there isn't lost due to React state batching.
      const matchingRow = prev.rows.find(r =>
        r.uploadFolderId === uploadFolderId ||
        r.uploadFolderId === resolvedId ||
        (groupKey && r.groupKey === groupKey)
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
        ...(matchingRow && matchingRow.subFolderPath ? (prev.uploadedFilesByFolder[(matchingRow.subFolderPath || '').trim().toLowerCase()] || []) : []),
        ...(matchingRow ? (prev.uploadedFilesByFolder[(matchingRow.subFolderPath || matchingRow.groupKey).trim().toLowerCase()] || []) : []),
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
      if (isUploadFolderIdGraph) updatedByFolder[uploadFolderId] = mergedUploads;
      if (isResolvedIdGraph) updatedByFolder[resolvedId] = mergedUploads;

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
        // Do NOT write to baseRow0.uploadFolderId if it's a backend DB ID
        if (baseRow0.uploadFolderId && this.state.spoFolderMap.has(baseRow0.uploadFolderId)) {
          updatedByFolder[baseRow0.uploadFolderId] = mergedUploads;
        }
      }

      let newRows: FlatRow[] = prev.rows;

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
          return { rows: prev.rows, uploadedFilesByFolder: safeByFolder };
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
      return { rows: newRows, uploadedFilesByFolder: updatedByFolder };
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

    return (
      <div style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 100001,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}>
        <div style={{
          background: '#fff', borderRadius: 16, padding: '28px 32px', minWidth: 420, maxWidth: 540,
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

          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
            <button
              onClick={() => this.setState({ fileDeleteDialog: null })}
              disabled={busy}
              style={{ padding: '8px 20px', borderRadius: 8, border: '1px solid #cbd5e1', background: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', color: '#334155' }}
            >
              Cancel
            </button>
            <button
              onClick={() => void this._deleteSelectedFiles()}
              disabled={busy || selected.size === 0}
              style={{
                padding: '8px 20px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 600, cursor: busy || selected.size === 0 ? 'not-allowed' : 'pointer',
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

  public _handleUpload = async (row: GroupedRow, files: File[]): Promise<void> => {
    if (!files.length) return;
    this.setState({ uploadingGroupKey: row.groupKey, uploadError: null, uploadInfo: `Uploading ${files.length === 1 ? files[0].name : `${files.length} files`}…` });
    let done = 0, failed = 0, pending = 0;
    let lastFolderId: string | null = null;
    let lastIsGraphUpload = false;
    let lastWebUrl = '';
    let lastDestPath = row.subFolderPath || row.subCategory || row.category;
    for (const file of files) {
      try {
        const result = await this._uploadFileToFolder(
          row.uploadFolderId,
          row.subFolderPath,
          row.vesselName,
          file,
          row.monthDriven,
        );
        if (result.statusPending) pending++;
        else done++;
        lastFolderId = result.folderId || lastFolderId;
        lastIsGraphUpload = result.isGraphUpload || lastIsGraphUpload;
      } catch (error) {
        failed++;
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
      await this._refreshFolderFiles(refreshFolderId, row.groupKey, true, isRealGraphId);
      // Trigger delta sync immediately so SPO changes propagate
      void this._syncScheduler?.triggerNow().catch(() => undefined);
    }
  };

  public _renderUploadSuccessPopup(): React.ReactElement | null {
    const p = this.state.uploadSuccessPopup;
    if (!p) return null;
    return (
      <div style={{
        position: 'fixed', bottom: 28, right: 28, zIndex: 100002,
        background: '#fff', borderRadius: 16, boxShadow: '0 8px 40px rgba(0,0,0,0.18)',
        border: '1.5px solid #86efac', padding: '20px 24px 18px', minWidth: 340, maxWidth: 420,
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
            style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#94a3b8', fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}
            title="Close"
          >✕</button>
        </div>
        {/* File info */}
        <div style={{ background: '#f8fafc', borderRadius: 8, padding: '8px 12px', marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: '#64748b', marginBottom: 3 }}>📄 File</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', wordBreak: 'break-all' }}>{p.fileName}</div>
          <div style={{ fontSize: 11, color: '#64748b', marginTop: 6 }}>📁 Path</div>
          <div style={{ fontSize: 12, color: '#334155', marginTop: 2, wordBreak: 'break-all' }}>{p.destinationPath}</div>
        </div>
        {/* Action buttons */}
        <div style={{ display: 'flex', gap: 8 }}>
          {p.webUrl && (
            <a
              href={p.webUrl}
              target="_blank"
              rel="noopener noreferrer"
              style={{
                flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                background: '#0078d4', color: '#fff', borderRadius: 8,
                padding: '8px 12px', fontSize: 12, fontWeight: 600, textDecoration: 'none',
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
              padding: '8px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            Close
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
    return renderDocumentsPage(this);
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
   * Restore a vessel folder from the SPO site Recycle Bin back to its original
   * Vessels/Specific Vessels/{name} location using the SharePoint REST API.
   * Returns true if the item was found and restored.
   */
  private async _restoreSpoVesselFolder(vesselName: string): Promise<boolean> {
    if (!this.props.siteUrl) return false;
    try {
      // Find the vessel folder in the SPO recycle bin by its leaf name
      const rbRes = await fetch(
        `${this.props.siteUrl}/_api/site/RecycleBin?$filter=LeafName eq '${encodeURIComponent(vesselName)}'&$select=Id,LeafName,DirName&$top=10`,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!rbRes.ok) return false;
      const rbData = await rbRes.json();
      const match = (rbData?.value ?? []).find(
        (r: any) => (r.LeafName || '').toLowerCase() === vesselName.toLowerCase()
      );
      if (!match?.Id) return false;

      // Get request digest for CSRF
      const digestRes = await fetch(`${this.props.siteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: { Accept: 'application/json;odata=nometadata' },
      });
      const digestData = await digestRes.json();
      const digest = digestData?.FormDigestValue ?? '';

      // Restore the item from the SPO recycle bin
      const restoreRes = await fetch(
        `${this.props.siteUrl}/_api/site/RecycleBin('${match.Id}')/restore()`,
        {
          method: 'POST',
          headers: {
            Accept: 'application/json;odata=nometadata',
            'X-RequestDigest': digest,
          },
        }
      );
      return restoreRes.ok || restoreRes.status === 204;
    } catch (err) {
      console.warn('[VesselDMS] _restoreSpoVesselFolder failed:', err);
      return false;
    }
  }

  public _restoreFromRecycleBin = async (item: DeletedNode): Promise<void> => {
    const isVessel = item.kind === 'vessel' || item.item_type === 'vessel';

    // For vessels: first try to restore the SPO folder from the SharePoint Recycle Bin
    // so the folder reappears in Vessels/Specific Vessels/{name} before we re-create the DB record.
    if (isVessel) {
      const spoRestored = await this._restoreSpoVesselFolder(item.name);
      if (!spoRestored) {
        // SPO restore failed or folder not in recycle bin — fall back to re-provisioning
        console.warn(`[VesselDMS] SPO recycle bin restore failed for "${item.name}", will re-provision folders after DB restore.`);
      }
    }

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
      if (!res.ok && res.status !== 202) throw new Error(`Restore failed (${res.status})`);
      if (res.status === 202) {
        window.alert(`Restore request for "${item.name}" was submitted for approval.`);
        return;
      }
    } catch (e: any) {
      window.alert(e?.message || `Could not restore "${item.name}". Please refresh the Recycle Bin and try again.`);
      return;
    }

    // If it's a vessel, restore vessel record into active vessels state
    if (isVessel) {
      const base = this._base();
      let restoredId = item.id.replace(/^vessel_/, '');
      try {
        const created = await this._fetchJson(`${base}/api/vessels`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: item.name,
            imo: item.imo && item.imo !== '—' ? item.imo : null,
            shipyard: item.shipyard || 'Restored',
            vessel_type: item.vessel_type || 'Bulk Carrier',
          }),
        });
        if (created?.id) restoredId = created.id;
      } catch { /* already exists or offline fallback */ }

      const restoredVessel: VesselRecord = {
        id: restoredId,
        name: cleanName(item.name),
        imo: item.imo || '—',
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

    this.setState(prev => ({ recycleBin: prev.recycleBin.filter(r => r.id !== item.id) }));
    // Trigger delta sync so the restored SPO folder appears in the tree
    void this._syncScheduler?.triggerNow().catch(() => undefined);
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
    return (
      <div style={{
        position: 'fixed', bottom: 28, left: '50%', transform: 'translateX(-50%)',
        zIndex: 100003, background: '#1e293b', color: '#fff',
        borderRadius: 12, boxShadow: '0 8px 32px rgba(0,0,0,0.28)',
        padding: '16px 24px', display: 'flex', alignItems: 'center', gap: 16,
        minWidth: 360, maxWidth: 520, fontFamily: "'Segoe UI', sans-serif",
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
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, flexShrink: 0 }}>
          <button
            onClick={() => {
              this.setState({ spoVesselDeletedToast: null });
              void this._goToView('recycle');
            }}
            style={{
              background: '#ef4444', color: '#fff', border: 'none', borderRadius: 7,
              padding: '5px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
            }}
          >View Recycle Bin</button>
          <button
            onClick={() => this.setState({ spoVesselDeletedToast: null })}
            style={{
              background: 'transparent', color: '#94a3b8', border: '1px solid #475569',
              borderRadius: 7, padding: '5px 12px', fontSize: 11, fontWeight: 600, cursor: 'pointer',
            }}
          >Dismiss</button>
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
      default:
        content = this._renderDashboard();
    }

    return (
      <>
        {this._renderLayout(content)}
        {this._renderBentoComposeModal()}
        {this._renderUploadSuccessPopup()}
        {this._renderFileDeleteDialog()}
        {this._renderSpoVesselDeletedToast()}
        {renderFileAlertDialog(this)}
      </>
    );
  }
}
