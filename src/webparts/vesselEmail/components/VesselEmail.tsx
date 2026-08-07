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
  ApprovalItem, NotificationItem, UserItem,
} from './types/ui';
import {
  cleanName, INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES,
  INITIAL_MOCK_NOTIFICATIONS, INITIAL_MOCK_USERS,
} from './constants';
import { renderSidebar } from './pages/Sidebar';
import { renderLayout } from './pages/AppLayout';
import { renderDocPreviewDrawer } from './pages/DocPreviewDrawer';
import { renderDashboard } from './pages/DashboardPage';
import { renderDocumentsPage } from './pages/DocumentsPage';
import { renderVesselsPage } from './pages/VesselsPage';
import { renderTemplatesPage } from './pages/TemplatesPage';
import { renderApprovalsPage } from './pages/ApprovalsPage';
import { renderNotificationsPage } from './pages/NotificationsPage';
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
  form: FormState;
  modalBusy: boolean;
  modalMsg: string | null;
  modalError: string | null;

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
  docUploadRowKey: string | null;
  docUploadBusy: boolean;
  docUploadMsg: string | null;
  folderPathStack: { id: string; name: string }[];
  uploadedFilesByFolder: Record<string, { name: string; size: string; date: string; pending?: boolean }[]>;
  selectedDocPreview: DocPreviewItem | null;
  templatesList: TemplateItem[];
  approvalsList: ApprovalItem[];
  approvalTab: 'Pending' | 'Approved' | 'Rejected';
  notificationsList: NotificationItem[];
  notificationFilter: 'all' | 'unread';
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

  // Delta sync — flat id→node map representing the live SPO folder tree
  spoFolderMap: Map<string, SpoFolderNode>;
  lastDeltaSync: Date | null;

  sessionExpired: boolean;
  sessionReady: boolean;  // true once first valid session_id prop is received
}

const BLANK_FORM: FormState = { name: '', imo: '', shipyard: '', hull_number: '', vessel_type: '' };

const PAGE_SIZE = 50;

// ── Component Definition ────────────────────────────────────────────────────

export default class VesselEmail extends React.Component<IVesselEmailProps, State> {
  public _abort: AbortController | null = null;
  public _filesLoadedForVessels: Set<string> = new Set();
  public _syncScheduler: SyncScheduler | null = null;

  public constructor(props: IVesselEmailProps) {
    super(props);
    this.state = {
      rows: [],
      vessels: [],
      loading: false, error: null, reloadKey: 0,
      textFilter: '', vesselFilter: 'all', groupFilter: 'all', catFilter: 'all',
      sort: 'default', uploadingGroupKey: null, uploadInfo: null, uploadError: null,
      selectedFileIds: new Set(), page: 0,
      modal: 'none', selectedVessel: null, form: { ...BLANK_FORM },
      modalBusy: false, modalMsg: null, modalError: null,
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
      docUploadRowKey: null,
      docUploadBusy: false,
      docUploadMsg: null,
      folderPathStack: [],
      uploadedFilesByFolder: {},
      selectedDocPreview: null,
      templatesList: INITIAL_MOCK_TEMPLATES,
      approvalsList: [],
      approvalTab: 'Pending',
      notificationsList: INITIAL_MOCK_NOTIFICATIONS,
      notificationFilter: 'all',
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
      spoFolderMap: new Map(),
      lastDeltaSync: null,
      sessionExpired: false,
      sessionReady: false,
    };
  }

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
  }

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
    // On-demand file loading when user selects a specific vessel in the Documents list view
    const { vesselFilter, rows } = this.state;
    if (ps.vesselFilter !== vesselFilter && vesselFilter !== 'all' && rows.length > 0) {
      if (!this._filesLoadedForVessels.has(vesselFilter)) {
        this._filesLoadedForVessels.add(vesselFilter);
        this._loadFilesForVessel(vesselFilter).catch(() => undefined);
      }
    }
  }

  public componentWillUnmount(): void {
    this._abort?.abort();
    this._syncScheduler?.stop();
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
      45_000,
      (err) => console.warn('[VesselDMS] Delta sync error:', err),
    );
    this._syncScheduler.start();
    // Run immediately on mount
    void this._syncScheduler.triggerNow().catch(() => undefined);
  }

  public _applyDeltaResult(result: DeltaSyncResult): void {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return;

    this.setState(prev => {
      const map = new Map(prev.spoFolderMap);

      // Process deletions first
      for (const id of result.deleted) {
        removeNodeFromMap(map, id);
      }

      // Process additions/updates
      const missingParentIds = new Set<string>();
      for (const node of result.added) {
        const { missingParentId } = mergeNodeIntoMap(map, node);
        if (missingParentId) missingParentIds.add(missingParentId);
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

      return { spoFolderMap: map, lastDeltaSync: new Date() };
    });
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

  public _fetchJson(url: string, signal?: AbortSignal): Promise<any> {
    return fetch(url, { signal, headers: this._headers() })
      .then(r => {
        if (r.status === 401 && this.state.sessionReady) {
          // Only show expired screen if session was previously confirmed valid
          this.setState({ sessionExpired: true });
          throw new Error('SESSION_EXPIRED');
        }
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.json();
      });
  }

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
    return url || 'http://localhost:8000';
  }

  // ── Data Loading ──────────────────────────────────────────────────────────

  public _loadData(): void {
    // NOTE: we intentionally do NOT require this.props.sessionId here.
    // A backend running without a database (settings.db_configured === false)
    // legitimately never issues a session_id, and require_session() on the
    // server allows unauthenticated calls in that mode. Gating on sessionId
    // alone left the vessel list permanently blank in that configuration.
    // _headers() already omits the Authorization/X-Session-ID headers when
    // sessionId is empty, so this is safe to call regardless.
    this._abort?.abort();
    this._abort = new AbortController();
    const signal = this._abort.signal;
    const base = this._base();
    if (!base) {
      this.setState({ loading: false, rows: [] });
      return;
    }
    this.setState({ loading: true, error: null });

    // Step 1: Always fetch vessel list from backend database first
    this._fetchJson(`${base}/api/vessels`, signal)
      .catch(() => null)
      .then(async (vesselList: any) => {
        if (signal.aborted) return;

        let vessels: VesselRecord[] = [];
        if (vesselList && Array.isArray(vesselList) && vesselList.length > 0) {
          vessels = vesselList.map((v: any) => ({
            ...v,
            name: cleanName(v.name),
            status: v.status || 'Active',
          }));
        }

        // Update vessel list in state immediately so UI shows vessels right away
        if (vessels.length > 0) {
          this.setState({ vessels });
        }

        // Step 2: Try Graph API folder-walk first (when SPO context is available)
        const { graphClient, siteId, driveId } = this.props;
        if (graphClient && siteId && driveId && vessels.length > 0) {
          console.log('[VesselDMS] _loadData: Graph context available — using Graph API for folder tree walk');
          try {
            const graphRows = await this._flattenAllViaGraph(vessels, signal);
            if (!signal.aborted) {
              console.log(`[VesselDMS] _loadData: Graph walk returned ${graphRows.length} rows`);
              this.setState({ vessels, rows: this._normalize(graphRows), loading: false });
              return;
            }
          } catch (graphErr) {
            console.warn('[VesselDMS] _loadData: Graph walk failed, falling back to REST API:', graphErr);
          }
        }

        if (signal.aborted) return;

        // Step 3: Fallback — try backend flat-tree endpoint
        const flatTree = await this._fetchJson(`${base}/api/vessels/flat-tree`, signal).catch(() => null);
        if (!signal.aborted) {
          if (flatTree && Array.isArray(flatTree) && flatTree.length > 0) {
            this.setState({ vessels, rows: this._normalize(flatTree), loading: false });
          } else if (vessels.length > 0) {
            // Step 4: Final fallback — REST API per-vessel/per-folder walk
            const rows = await this._flattenAll(vessels, signal);
            if (!signal.aborted) {
              this.setState({ vessels, rows: this._normalize(rows), loading: false });
            }
          } else {
            this.setState({ vessels, rows: [], loading: false });
          }
        }
      })
      .catch(err => {
        if (!signal.aborted) this.setState({ loading: false, error: err?.message ?? 'Failed to load data.' });
      });
  }

  public _normalize(raw: FlatRow[]): FlatRow[] {
    return raw.map(r => ({ ...r, vesselName: cleanName(r.vesselName), group: cleanName(r.group), category: cleanName(r.category) }));
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
  ): Promise<{ fileId: string | null; statusPending: boolean }> {
    const base = this._base();
    if (!base) return { fileId: null, statusPending: false };

    const form = new FormData();
    form.append('file', file);
    if (this.props.userEmail) form.append('uploader_email', this.props.userEmail);

    let endpoint: string;
    if (uploadFolderId.includes('/')) {
      endpoint = `${base}/api/folders/upload-by-path?path=${encodeURIComponent(uploadFolderId)}`;
    } else {
      endpoint = monthDriven
        ? `${base}/api/folders/${encodeURIComponent(uploadFolderId)}/month-upload`
        : `${base}/api/folders/${encodeURIComponent(uploadFolderId)}/upload`;
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
    return { fileId: data?.id || null, statusPending: data?.status === 'pending' };
  }

  // ── Graph API Folder Walking (mirrors VesselListView.tsx from reference project) ────

  public readonly GRAPH_FETCH_CONCURRENCY = 6;
  public readonly VESSEL_ROOT = 'Vessel Management';
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
  ): Promise<Array<{ id: string; name: string; isFolder: boolean; monthDriven: boolean; upload: boolean }>> {
    const { graphClient, siteId, driveId } = this.props;
    if (!graphClient || !siteId || !driveId) return [];

    const encodedPath = folderPath.split('/').map(s => encodeURIComponent(s)).join('/');
    const url = `/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/children` +
      `?$select=id,name,folder,file&$top=200`;

    console.log(`[VesselDMS] _getGraphChildren → ${url}`);
    try {
      const result: any = await graphClient.api(url).get();
      if (signal.aborted) return [];
      return (result.value ?? []).map((item: any) => ({
        id: item.id as string,
        name: item.name as string,
        isFolder: !!item.folder,
        monthDriven: false,   // Graph doesn't expose month_driven; backend knows this
        upload: !!item.folder, // any folder can receive uploads
      }));
    } catch (err: any) {
      // 404 means the folder doesn't exist yet (vessel not provisioned) — not an error
      const status = err?.statusCode ?? err?.code ?? 0;
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

    let kids: Array<{ id: string; name: string; isFolder: boolean; monthDriven: boolean; upload: boolean }>;
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
    // pathParts[0] = main folder name, pathParts[1] = first category (e.g. "Crewing")
    // GROUP should be the first category under the vessel (pathParts[1]), not the main folder
    const group = pathParts.length >= 2 ? stripPrefix(pathParts[1]) : stripPrefix(pathParts[0] ?? folderPath.split('/').pop() ?? folderPath);
    const category = pathParts.length >= 3 ? stripPrefix(pathParts[pathParts.length - 1]) : group;
    // subFolderPath: VesselName > CategoryName (skip main folder prefix)
    const subPath = pathParts.length >= 2
      ? [vesselName, ...pathParts.slice(1).map(stripPrefix)].join(' > ')
      : [vesselName, ...pathParts.map(stripPrefix)].join(' > ');
    const groupKey = `${vesselName}||${group}||${category}||${subPath}`;
    const canUpload = subFolders.length === 0; // leaf nodes can receive uploads

    // Emit a leaf row (matches reference leafToRows pattern)
    if (files.length > 0 || subFolders.length === 0) {
      srCounter.value += 1;
      const baseSr = String(srCounter.value);
      const suffixes = 'abcdefghijklmnopqrstuvwxyz';

      if (files.length === 0) {
        onRows([{
          srNo: baseSr, vesselName, group, category,
          subFolderPath: subPath, fileName: null, fileId: null,
          canUpload, groupKey,
          uploadFolderId: folderPath,
          monthDriven: false,
        }]);
      } else {
        const leafRows: FlatRow[] = files.map((f, idx) => ({
          srNo: idx === 0 ? baseSr : `${baseSr}${suffixes[idx - 1] ?? idx}`,
          vesselName, group, category, subFolderPath: subPath,
          fileName: f.name, fileId: f.id, canUpload, groupKey,
          uploadFolderId: folderPath,
          monthDriven: false,
        }));
        onRows(leafRows);
      }
    }

    // Recurse into subfolders with bounded concurrency
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

    // Walk each main folder in parallel
    await this._mapLimit(
      this.MAIN_FOLDER_NAMES,
      3,
      async mainFolderName => {
        if (signal.aborted) return;
        const vesselPath = `${this.VESSEL_ROOT}/${mainFolderName}/${vesselName}`;
        // Get top-level categories inside this vessel's main folder
        const topCats = await this._getGraphChildren(vesselPath, signal).catch(() => []);
        if (signal.aborted) return;

        const folders = topCats.filter(c => c.isFolder);
        const files = topCats.filter(c => !c.isFolder);

        if (folders.length === 0) {
          // Vessel folder is itself a leaf
          const displayMainFolder = mainFolderName.replace(/^Folder-\d+\s+/i, '');
          srCounter.value += 1;
          const leafRows: FlatRow[] = files.map((f, idx) => ({
            srNo: idx === 0 ? String(srCounter.value) : `${srCounter.value}${String.fromCharCode(97 + idx - 1)}`,
            vesselName,
            group: displayMainFolder,
            category: displayMainFolder,
            subFolderPath: `${vesselName} > ${displayMainFolder}`,
            fileName: f.name,
            fileId: f.id,
            canUpload: true,
            groupKey: `${vesselName}||${displayMainFolder}||${displayMainFolder}||${vesselName} > ${displayMainFolder}`,
            uploadFolderId: f.id,
            monthDriven: false,
          }));
          if (leafRows.length === 0) {
            leafRows.push({
              srNo: String(srCounter.value),
              vesselName,
              group: displayMainFolder,
              category: displayMainFolder,
              subFolderPath: `${vesselName} > ${displayMainFolder}`,
              fileName: null, fileId: null, canUpload: true,
              groupKey: `${vesselName}||${displayMainFolder}||${displayMainFolder}`,
              uploadFolderId: vesselPath,
              monthDriven: false,
            });
          }
          allRows.push(...leafRows);
          onChunk(leafRows);
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
              [mainFolderName, cat.name],
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
    let flushTimer: ReturnType<typeof setTimeout> | null = null;

    const queueFlush = (): void => {
      if (flushTimer !== null) return;
      flushTimer = setTimeout(() => {
        flushTimer = null;
        if (signal.aborted) return;
        const snapshot = [...allRows];
        this.setState(prev => ({ rows: this._normalize(snapshot), vessels: prev.vessels }));
      }, 120);
    };

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
            queueFlush();
          },
        );
      },
    );

    if (flushTimer !== null) {
      clearTimeout(flushTimer);
      flushTimer = null;
    }

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
          try {
            const cats = await this._fetchJson(`${base}/api/folders/${m.id}/children`, signal);
            for (const cat of cats) {
              if (cat?.kind === 'file') continue; sr++;
              const sp = `${v.name} > ${m.name} > ${cat.name}`;
              out.push({ srNo: String(sr), vesselName: v.name, group: m.name, category: cat.name, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: `${v.id}||${m.name}||${cat.name}||${sp}`, uploadFolderId: cat.id, monthDriven: Boolean(cat.month_driven) });
            }
          } catch { /* skip */ }
        }
      } catch { /* skip */ }
    }
    return out;
  }

  // ── Navigation & Views ────────────────────────────────────────────────────

  public _goToView = async (view: AppView): Promise<void> => {
    this.setState({ view });
    if (view === 'vessels') {
      this.setState({ panelLoading: true });
      try {
        const data = await this._fetchJson(`${this._base()}/api/vessels`);
        if (data && Array.isArray(data)) {
          this.setState({
            vessels: data.map((v: any) => ({
              ...v,
              name: cleanName(v.name),
              status: v.status || 'Active',
            })),
            panelLoading: false,
          });
        } else {
          // If REST API returned nothing, vessels already in state are still valid
          this.setState({ panelLoading: false });
        }
      } catch {
        this.setState({ panelLoading: false });
      }
    } else if (view === 'recycle') {
      this.setState({ panelLoading: true, recycleBin: [] });
      try {
        const data = await this._fetchJson(`${this._base()}/api/recycle-bin/nodes`);
        this.setState({ recycleBin: (data || []).map((v: any) => ({ ...v, name: cleanName(v.name || '') })), panelLoading: false });
      } catch {
        this.setState({ recycleBin: [], panelLoading: false });
      }
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
    } else if (view === 'notifications') {
      this.setState({ panelLoading: true });
      try {
        const base = this._base();
        const data = await this._fetchJson(`${base}/api/approvals?admin=${encodeURIComponent(this.props.userEmail || '')}`).catch(() => null);
        const myData = await this._fetchJson(`${base}/api/my-approvals`).catch(() => null);
        const combined: any[] = [];
        const seen = new Set<string>();
        for (const a of [...(Array.isArray(data) ? data : []), ...(Array.isArray(myData) ? myData : [])]) {
          const key = String(a.id || '');
          if (!seen.has(key)) { seen.add(key); combined.push(a); }
        }
        if (combined.length > 0) {
          const notifications: NotificationItem[] = combined.map((a: any) => {
            const fname = a.filename || a.file_name || a.name || 'Document';
            const vessel = a.vessel_name || a.vesselName || '';
            const uploader = a.uploaded_by_email || a.uploaded_by_name || '';
            const raw = a.uploaded_at || a.created_at || '';
            const ts = raw ? new Date(raw).toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
            const status = a.status || 'pending';
            const type: NotificationItem['type'] = status === 'approved' ? 'success' : status === 'rejected' ? 'alert' : 'warning';
            const priority: NotificationItem['priority'] = status === 'pending' ? 'High' : 'Low';
            const title = status === 'approved' ? 'Document Approved' : status === 'rejected' ? 'Document Rejected' : 'Approval Requested';
            const message = status === 'approved'
              ? `Document approved: ${fname}${vessel ? ` for ${vessel}` : ''}`
              : status === 'rejected'
              ? `Document rejected: ${fname}${vessel ? ` for ${vessel}` : ''}`
              : `Approval requested: ${fname}${vessel ? ` for ${vessel}` : ''}${uploader ? ` by ${uploader}` : ''}`;
            return { id: String(a.id), title, message, timestamp: ts, priority, read: status !== 'pending', type };
          });
          this.setState({ notificationsList: notifications, panelLoading: false });
        } else {
          this.setState({ panelLoading: false });
        }
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
    this.setState({ modal: 'create', selectedVessel: null, form: { ...BLANK_FORM }, modalMsg: null, modalError: null });
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

  public _openDeleteVessel = (v: VesselRecord): void => {
    this.setState({ modal: 'delete', selectedVessel: v, modalMsg: null, modalError: null });
  };

  public _closeModal = (): void => {
    if (!this.state.modalBusy) this.setState({ modal: 'none', modalMsg: null, modalError: null });
  };

  public _submitCreate = async (): Promise<void> => {
    const { form, vessels } = this.state;
    if (!form.name.trim()) { this.setState({ modalError: 'Vessel name is required.' }); return; }
    if (!form.imo.trim()) { this.setState({ modalError: 'IMO number is required.' }); return; }
    if (!/^\d{7}$/.test(form.imo.trim())) { this.setState({ modalError: 'IMO number must be exactly 7 digits.' }); return; }
    this.setState({ modalBusy: true, modalError: null, modalMsg: null });

    try {
      const res = await fetch(`${this._base()}/api/vessels`, {
        method: 'POST', headers: this._headers(),
        body: JSON.stringify({ name: form.name.trim(), imo: form.imo.trim(), shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null }),
      });
      const data = await res.json();
      if (!res.ok && res.status !== 202) throw new Error(data?.message ?? `Error ${res.status}`);
      const msg = data.status === 'pending' ? '⏳ Vessel creation submitted for approval.' : `✅ Vessel "${form.name}" created successfully.`;
      // Optimistically add the new vessel to the list so it shows immediately
      const newVesselRecord: VesselRecord = {
        id: data.id || data.result?.id || `v_${Date.now()}`,
        name: form.name.trim(),
        imo: form.imo.trim(),
        shipyard: form.shipyard.trim() || undefined,
        hull_number: form.hull_number.trim() || undefined,
        vessel_type: form.vessel_type || undefined,
        status: 'Active',
        image_url: pickRandomVesselImage(form.vessel_type),
      };
      this.setState({ modalBusy: false, modalMsg: msg, modalError: null, vessels: [...this.state.vessels, newVesselRecord] });
      // Kick off SharePoint folder provisioning in the background (non-blocking)
      this._provisionVesselFolders(form.name.trim()).catch(() => undefined);
      setTimeout(() => this.setState({ modal: 'none', reloadKey: this.state.reloadKey + 1 }), 1600);
    } catch (e: any) {
      // Fallback local creation if backend not available.
      // Pick a random image ONCE at creation time and persist it on the record.
      const newVessel: VesselRecord = {
        id: `v_${Date.now()}`,
        name: form.name.trim(),
        imo: form.imo.trim(),
        shipyard: form.shipyard.trim(),
        hull_number: form.hull_number.trim(),
        vessel_type: form.vessel_type || 'Bulk Carrier',
        status: 'Active',
        image_url: pickRandomVesselImage(form.vessel_type),
      };
      this.setState({
        vessels: [...vessels, newVessel],
        modalBusy: false,
        modalMsg: `✅ Vessel "${form.name}" created successfully.`,
      });
      // Kick off SharePoint folder provisioning in the background (non-blocking)
      this._provisionVesselFolders(form.name.trim()).catch(() => undefined);
      setTimeout(() => this.setState({ modal: 'none' }), 1200);
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
    const { selectedVessel, vessels } = this.state;
    if (!selectedVessel) return;
    this.setState({ modalBusy: true, modalError: null });
    try {
      const res = await fetch(`${this._base()}/api/vessels/${selectedVessel.id}?vessel_name=${encodeURIComponent(selectedVessel.name)}`, {
        method: 'DELETE', headers: this._headers(),
      });
      const data = await res.json();
      if (!res.ok && res.status !== 202) throw new Error(data?.message ?? `Error ${res.status}`);
      const updated = vessels.filter(v => v.id !== selectedVessel.id);
      const msg = data.status === 'pending'
        ? `Delete submitted for approval. "${selectedVessel.name}" removed from list.`
        : `"${selectedVessel.name}" moved to Recycle Bin.`;
      this.setState({ modalBusy: false, modalMsg: msg, vessels: updated });
      setTimeout(() => this.setState({ modal: 'none', selectedVessel: null }), 1800);
    } catch (e: any) {
      const updated = vessels.filter(v => v.id !== selectedVessel.id);
      this.setState({ vessels: updated, modalBusy: false, modalMsg: `"${selectedVessel.name}" deleted.` });
      setTimeout(() => this.setState({ modal: 'none', selectedVessel: null }), 1200);
    }
  };

  // SharePoint Folder Provisioning ─────────────────────────────────────────

  public async _provisionVesselFolders(vesselName: string, vesselId?: string): Promise<void> {
    const { graphClient, siteId, driveId } = this.props;
    console.log(`[VesselDMS] _provisionVesselFolders vessel="${vesselName}" graphClient=${!!graphClient} siteId="${siteId}" driveId="${driveId}"`);
    if (!graphClient || !siteId || !driveId) {
      console.error('[VesselDMS] _provisionVesselFolders aborted — missing graphClient/siteId/driveId. Props:', { graphClient: !!graphClient, siteId, driveId });
      this.setState({ folderCreationError: 'SharePoint context not ready. Please refresh the page and try again.' });
      return;
    }

    this.setState({ folderCreationBusy: true, folderCreationResults: null, folderCreationError: null, folderProvisioningVesselId: vesselId || null });
    try {
      const result = await createVesselFolders(graphClient, siteId, driveId, vesselName);
      this.setState({
        folderCreationBusy: false,
        folderProvisioningVesselId: null,
        folderCreationResults: result.results,
        folderCreationError: result.success ? null : 'Some folders could not be created. Check the creation log.',
      });
      // Trigger an immediate delta sync so the new folders appear in the tree
      void this._syncScheduler?.triggerNow().catch(() => undefined);
    } catch (err: any) {
      this.setState({
        folderCreationBusy: false,
        folderProvisioningVesselId: null,
        folderCreationError: err?.message ?? 'Folder provisioning failed.',
      });
    }
  }

  // ── Load Files for Vessel ─────────────────────────────────────────────────

  public async _loadFilesForVessel(vesselName: string): Promise<void> {
    const { rows } = this.state;
    const vesselRows = rows.filter(r => r.vesselName === vesselName && r.uploadFolderId);
    const uniqueFolderIds = Array.from(new Set(vesselRows.map(r => r.uploadFolderId)));
    if (uniqueFolderIds.length === 0) return;

    const BATCH = 5;
    for (let i = 0; i < uniqueFolderIds.length; i += BATCH) {
      const batch = uniqueFolderIds.slice(i, i + BATCH);
      await Promise.all(batch.map(fid => this._refreshFolderFiles(fid, '')));
    }
  }

  public _filesLoadedForFolders: Set<string> = new Set();

  public async _refreshFolderFiles(uploadFolderId: string, groupKey: string, force: boolean = false): Promise<void> {
    if (!uploadFolderId) return;
    if (/^f\d+$/.test(uploadFolderId) || uploadFolderId.includes('/')) return;
    if (!force && this._filesLoadedForFolders.has(uploadFolderId)) return;
    this._filesLoadedForFolders.add(uploadFolderId);
    try {
      const children: any[] = await this._fetchJson(`${this._base()}/api/folders/${encodeURIComponent(uploadFolderId)}/children`);
      const fileItems = (children || []).filter((c: any) => c.kind === 'file');
      const parsedUploads = fileItems.map((f: any) => ({
        name: f.name || f.displayName,
        size: f.size ? `${(f.size / 1024).toFixed(1)} KB` : '142 KB',
        date: f.modified || 'Today',
        pending: false,
        id: f.id,
      }));

      this.setState(prev => {
        const updatedByFolder: Record<string, any[]> = { ...prev.uploadedFilesByFolder, [uploadFolderId]: parsedUploads };
        const existingBaseRows = prev.rows.filter(r => r.uploadFolderId === uploadFolderId);
        let newRows: FlatRow[] = prev.rows;

        if (existingBaseRows.length > 0) {
          const baseRow = existingBaseRows[0];
          updatedByFolder[baseRow.groupKey] = parsedUploads;
          updatedByFolder[baseRow.category] = parsedUploads;
          updatedByFolder[baseRow.group] = parsedUploads;

          const otherFolderRows = prev.rows.filter(r => r.uploadFolderId !== uploadFolderId);
          if (fileItems.length === 0) {
            newRows = [...otherFolderRows, { ...baseRow, fileName: null, fileId: null, filePending: false }];
          } else {
            const mappedRows = fileItems.map((f: any, i: number) => ({
              ...baseRow,
              srNo: i === 0 ? baseRow.srNo : `${baseRow.srNo}.${i + 1}`,
              fileName: f.name || f.displayName || null,
              fileId: f.id || null,
              filePending: false,
            }));
            newRows = [...otherFolderRows, ...mappedRows];
          }
        }
        return { rows: newRows, uploadedFilesByFolder: updatedByFolder };
      });
    } catch { /* ignore */ }
  }

  public _handleUpload = async (row: GroupedRow, files: File[]): Promise<void> => {
    if (!files.length) return;
    this.setState({ uploadingGroupKey: row.groupKey, uploadError: null, uploadInfo: `Uploading ${files.length === 1 ? files[0].name : `${files.length} files`}…` });
    let done = 0, failed = 0, pending = 0;
    for (const file of files) {
      try {
        const fd = new FormData(); fd.append('file', file);
        const endpoint = row.monthDriven
          ? `${this._base()}/api/folders/${row.uploadFolderId}/month-upload`
          : `${this._base()}/api/folders/${row.uploadFolderId}/upload`;
        const h = this._uploadHeaders();  // no Content-Type – browser sets multipart boundary
        const res = await fetch(endpoint, { method: 'POST', body: fd, headers: h });
        if (res.ok || res.status === 202) {
          try {
            const data = await res.json();
            if (data?.status === 'pending') { pending++; } else { done++; }
          } catch { done++; }
        } else { failed++; }
      } catch { failed++; }
    }
    let infoMsg: string | null = null;
    let errMsg: string | null = null;
    if (failed > 0) {
      errMsg = `❌ ${failed} file(s) failed. ${done} succeeded.`;
    } else if (pending > 0 && done === 0) {
      infoMsg = `⏳ ${pending} file(s) submitted for approval — status set to Pending.`;
    } else if (pending > 0) {
      infoMsg = `✅ ${done} uploaded. ⏳ ${pending} pending approval.`;
    } else {
      infoMsg = `✅ ${done} file(s) uploaded successfully.`;
    }
    this.setState({ uploadingGroupKey: null, uploadInfo: infoMsg, uploadError: errMsg });
    if (done > 0) {
      await this._refreshFolderFiles(row.uploadFolderId, row.groupKey, true);
    }
  };

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
      if (!ex) map.set(row.groupKey, { srNo: row.srNo, vesselName: row.vesselName, group: row.group, category: row.category, subFolderPath: row.subFolderPath, groupKey: row.groupKey, uploadFolderId: row.uploadFolderId, monthDriven: row.monthDriven, canUpload: row.canUpload, files: row.fileId && row.fileName ? [{ id: row.fileId, name: row.fileName }] : [] });
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

  public _renderNotificationsPage(): React.ReactElement {
    return renderNotificationsPage(this);
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
        await this._loadFilesForVessel(vesselName).catch(() => undefined);
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

  public _restoreFromRecycleBin = async (item: DeletedNode): Promise<void> => {
    try {
      await fetch(`${this._base()}/api/recycle-bin/nodes/${item.id}/restore`, {
        method: 'POST', headers: this._headers(),
      });
    } catch { /* fallback */ }
    this.setState(prev => ({ recycleBin: prev.recycleBin.filter(r => r.id !== item.id) }));
    if (item.kind === 'vessel' || item.item_type === 'vessel') {
      this.setState(prev => ({ reloadKey: prev.reloadKey + 1 }));
    }
  };

  public _permanentDeleteFromRecycleBin = async (item: DeletedNode): Promise<void> => {
    if (!window.confirm(`Permanently delete "${item.name}"? This cannot be undone.`)) return;
    try {
      await fetch(`${this._base()}/api/recycle-bin/nodes/${item.id}`, {
        method: 'DELETE', headers: this._headers(),
      });
    } catch { /* fallback */ }
    this.setState(prev => ({ recycleBin: prev.recycleBin.filter(r => r.id !== item.id) }));
  };

  public _renderRecycleBinPage(): React.ReactElement {
    return renderRecycleBinPage(this);
  }

  public _renderArchivePage(): React.ReactElement {
    return renderArchivePage(this);
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
      case 'notifications':
        content = this._renderNotificationsPage();
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
      </>
    );
  }
}
