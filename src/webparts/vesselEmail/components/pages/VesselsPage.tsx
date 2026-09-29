/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/no-unescaped-entities */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import {
  badge, GROUP_COLORS, DATASOURCE_TAGS_MAP, VESSEL_TYPES, cleanName, suggestTagFromFilename,
  INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES, INITIAL_MOCK_APPROVALS,
  INITIAL_MOCK_NOTIFICATIONS, INITIAL_MOCK_USERS,
} from '../constants';
import type {
  FlatRow, GroupedRow, VesselRecord,
} from '../types/rows';
import type { BentoEmailLog } from '../types/bento';
import type { AppView, ModalMode } from '../types/view';
import type {
  FormState, DocPreviewItem, DeletedNode, DocumentItem, TemplateItem,
  ApprovalItem, NotificationItem, UserItem, FolderAnomalyItem, NormalFolderRecord,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import { KAIZEN_KNOWLEDGE_BANK_TREE, MAIN_FOLDERS } from '../vesselFolderTemplate';
import { isMobileWidth } from '../responsive';
import { clay } from '../clayTheme';


function getSpoVesselFolderUrl(
  hostOrSiteUrl?: VesselEmail | string,
  vesselOrName?: VesselRecord | string,
  siteUrlProp?: string
): string {
  let host: VesselEmail | undefined;
  let siteUrl = '';
  let vesselName = '';
  let vesselObj: VesselRecord | undefined;

  if (hostOrSiteUrl && typeof hostOrSiteUrl === 'object' && 'state' in hostOrSiteUrl) {
    host = hostOrSiteUrl as VesselEmail;
    if (vesselOrName && typeof vesselOrName === 'object' && 'name' in vesselOrName) {
      vesselObj = vesselOrName as VesselRecord;
      vesselName = vesselObj.name || '';
    } else if (typeof vesselOrName === 'string') {
      vesselName = vesselOrName;
    }
    const documentSites = host?.state?.documentSites || [];
    siteUrl = siteUrlProp || (vesselObj ? getVesselSharePointSiteUrl(host, vesselObj) : (documentSites.find(s => s.site_key === host?.state?.activeDocumentSite)?.web_url || host?.props?.siteUrl || ''));
  } else {
    siteUrl = typeof hostOrSiteUrl === 'string' ? hostOrSiteUrl : '';
    if (vesselOrName && typeof vesselOrName === 'object' && 'name' in vesselOrName) {
      vesselName = (vesselOrName as VesselRecord).name || '';
      vesselObj = vesselOrName as VesselRecord;
    } else if (typeof vesselOrName === 'string') {
      vesselName = vesselOrName;
    }
  }

  if (host && vesselName) {
    const cleanVessel = vesselName.trim().toLowerCase();
    const liveMatch = (host.state?.documentLiveFolders || []).find(f =>
      f.is_folder !== false && f.name.trim().toLowerCase() === cleanVessel
    );
    if (liveMatch?.web_url) {
      return liveMatch.web_url;
    }
  }

  const fallbackSite = 'https://nissenkaiunsingapore.sharepoint.com';
  const effectiveSite = (siteUrl && siteUrl !== '#') ? siteUrl.trim() : fallbackSite;

  try {
    const urlObj = new URL(effectiveSite);
    const basePath = urlObj.pathname.replace(/\/$/, '');

    let libName = 'Shared Documents';
    if (host) {
      const documentSites = host.state?.documentSites || [];
      const safeVesselObj = vesselObj;
      const safeHost = host;
      const matchedSite = safeVesselObj
        ? documentSites.find(s => {
            const vesselSiteIds = safeVesselObj.provisioned_site_ids || [];
            return vesselSiteIds.some(id => String(id).toLowerCase() === String(s.site_key || '').toLowerCase()) ||
              ((safeVesselObj.provisioned_site_key || '').toLowerCase() === String(s.site_key || '').toLowerCase()) ||
              s.web_url === effectiveSite;
          })
        : documentSites.find(s => s.site_key === safeHost.state?.activeDocumentSite || s.web_url === effectiveSite);
      const siteKey = String(matchedSite?.site_key || '').toLowerCase();
      const siteName = String(matchedSite?.sp_site_name || '').toLowerCase();
      const isCommunicationSite = siteKey === 'dev' || siteKey === 'communication' || siteKey === 'root' ||
        siteName === 'communication site' || siteName.includes('communication');
      if (isCommunicationSite) {
        // The tenant root Communication site uses the SharePoint library
        // named "Shared Documents", even when its configuration record says
        // "Documents" (the latter is the default for other site types).
        libName = 'Shared Documents';
      } else if (matchedSite?.default_library_name) {
        libName = matchedSite.default_library_name;
      }
    }

    const encodedLib = encodeURIComponent(libName);
    const customPath = vesselObj?.vessel_folder_path?.trim().replace(/^\/+|\/+$/g, '');
    if (customPath) {
      const pathParts = customPath.split('/').map(p => p.trim()).filter(Boolean);
      const normalizedParts = pathParts[0] && pathParts[0].toLowerCase() === libName.toLowerCase() ? pathParts.slice(1) : pathParts;
      const relativePath = normalizedParts.map(part => encodeURIComponent(part)).join('/');
      return `${urlObj.origin}${basePath}/${encodedLib}/${relativePath}`;
    }

    if (!vesselName) {
      return `${urlObj.origin}${basePath}/${encodedLib}`;
    }

    let deptName = 'Technical & Crewing';
    if (host) {
      const liveDept = (host.state.documentLiveFolders || []).find(f =>
        f.depth === 0 && /technical/i.test(f.name)
      );
      if (liveDept?.name) {
        deptName = liveDept.name;
      }
    }

    const encodedDept = encodeURIComponent(deptName);
    const encodedVessel = encodeURIComponent(vesselName.trim());
    return `${urlObj.origin}${basePath}/${encodedLib}/${encodedDept}/${encodedVessel}`;
  } catch {
    return '#';
  }
}

function getVesselSharePointSiteUrl(host: VesselEmail, vessel: VesselRecord): string {
  const candidateSiteIds = Array.from(new Set([...(vessel.provisioned_site_ids || []), vessel.provisioned_site_key].filter(Boolean) as string[]));

  // 1. Try matching with known document sites
  const targetSite = host.state.documentSites.find(site =>
    candidateSiteIds.some(id => {
      const idStr = String(id).toLowerCase();
      return (
        idStr === site.site_key.toLowerCase() ||
        (!!site.site_id && idStr === site.site_id.toLowerCase()) ||
        (!!site.sp_site_name && idStr === site.sp_site_name.toLowerCase())
      );
    })
  ) || (host.state.activeDocumentSite ? host.state.documentSites.find(s => s.site_key === host.state.activeDocumentSite) : null);

  if (targetSite?.web_url) {
    let url = targetSite.web_url.trim();
    // If target site is NKSDocMan but web_url was set without /sites/, fix it
    if ((targetSite.site_key.toLowerCase().includes('nks') || targetSite.sp_site_name.toLowerCase().includes('nks')) && !url.includes('/sites/')) {
      url = 'https://nissenkaiunsingapore.sharepoint.com/sites/NKSDocMan';
    }
    return url;
  }

  // 2. Fallback heuristic from provisionedSiteIds or activeDocumentSite
  const hasNks = candidateSiteIds.some(id => {
    const s = String(id).toLowerCase();
    return s.includes('nks') || s.includes('docman') || s === 'local';
  });
  if (hasNks || (host.state.activeDocumentSite && (host.state.activeDocumentSite.toLowerCase().includes('nks') || host.state.activeDocumentSite.toLowerCase() === 'local'))) {
    return 'https://nissenkaiunsingapore.sharepoint.com/sites/NKSDocMan';
  }

  const hasExternal = candidateSiteIds.some(id => String(id).toLowerCase().includes('external'));
  if (hasExternal || (host.state.activeDocumentSite && host.state.activeDocumentSite.toLowerCase().includes('external'))) {
    return 'https://nissenkaiunsingapore.sharepoint.com/sites/NissenKaiunExternal';
  }

  return host.props.siteUrl || 'https://nissenkaiunsingapore.sharepoint.com';
}

function resolveVesselDocumentNavigation(host: VesselEmail, vessel: VesselRecord): {
  docMainFolder: VesselEmail['state']['docMainFolder'];
  folderPathStack: { id: string; name: string }[];
  vesselFilter: string;
} {
  const vesselName = cleanName(vessel.name || '').trim();
  const customPath = (vessel.vessel_folder_path || '').trim();
  const fallbackMain = host.state.docMainFolder || 'Technical & Crewing';
  // Resolve the vessel's OWN site first (same candidate matching
  // getVesselSharePointSiteUrl already uses), before ever falling back to
  // whatever site happens to be active in the app right now. Falling back
  // to activeDocumentSite/documentSites[0] first (the previous order here)
  // meant "View Documents" ignored provisioned_site_key/provisioned_site_ids
  // entirely whenever the active site's own site_key satisfied that check
  // trivially — landing on the wrong site's (usually empty) folder for any
  // vessel whose real site differs from the currently active one, most
  // visibly the "Found in SharePoint" rows merged from a non-active site.
  const vesselSiteIds = Array.from(new Set(
    [...(vessel.provisioned_site_ids || []), vessel.provisioned_site_key].filter(Boolean) as string[]
  ));
  const selectedSite =
    host.state.documentSites.find(site => vesselSiteIds.some(id => {
      const idStr = String(id).toLowerCase();
      return idStr === (site.site_key || '').toLowerCase() ||
        (!!site.site_id && idStr === site.site_id.toLowerCase()) ||
        (!!site.sp_site_name && idStr === site.sp_site_name.toLowerCase());
    })) ||
    host.state.documentSites.find(site =>
      !!site.web_url && getVesselSharePointSiteUrl(host, vessel).startsWith(site.web_url) ||
      site.site_key === host.state.activeDocumentSite ||
      site.sp_site_name.toLowerCase() === (host.state.activeDocumentSite || '').toLowerCase()
    ) || host.state.documentSites[0];

  const rootStack: { id: string; name: string }[] = [];
  if (selectedSite) {
    // These three nodes (sites_root sentinel, then the site, then the
    // drive) are the exact shape DocumentsPage.tsx's live-folder browsing
    // expects at indices [0]/[1]/[2] (atSitesRoot, isSitesScopeNav, the
    // click handler below's own folderPathStack[1]/[0]==='sites_root'
    // checks). Without the 'sites_root' sentinel this stack was
    // off-by-one — folderPathStack[0] held the site node, so every one of
    // those checks failed, docScopeType always fell back to 'vessels'
    // instead of 'sites', and the live Graph children fetch that actually
    // populates Folder view never ran. That's what made "View Documents"
    // (and any breadcrumb click rebuilding a slice of this same stack)
    // land on a permanently empty listing for a flat, Part-C-provisioned
    // vessel folder.
    rootStack.push({ id: 'sites_root', name: 'SharePoint Sites' });
    rootStack.push({ id: `site:${selectedSite.site_id || selectedSite.site_key}`, name: selectedSite.sp_site_name || selectedSite.site_key || 'Communication Site' });
    rootStack.push({ id: `drive:${selectedSite.drive_id || 'root'}`, name: selectedSite.default_library_name || 'Documents' });
  }

  const fallbackDocMainFolder = (fallbackMain || 'Technical & Crewing') as VesselEmail['state']['docMainFolder'];

  if (!customPath) {
    const docMainFolder = fallbackDocMainFolder;
    return {
      docMainFolder,
      folderPathStack: [...rootStack, { id: String(vessel.id || vesselName), name: vesselName }],
      vesselFilter: vesselName,
    };
  }

  const rawSegments = customPath.replace(/\\/g, '/').split('/').map(part => part.trim()).filter(Boolean);
  const filteredSegments = rawSegments.filter(segment => !['SharePoint Sites', 'Documents', 'Communication Site', 'Shared Documents', 'Site Documents'].some(token => token.toLowerCase() === segment.toLowerCase()));
  const relativeSegments = filteredSegments;

  const mainMatch = relativeSegments.find(segment =>
    MAIN_FOLDERS.some((mf) => mf.name.toLowerCase() === segment.toLowerCase())
  );
  const resolvedMain = (mainMatch || fallbackDocMainFolder || 'Technical & Crewing') as string;
  const docMainFolder = resolvedMain as VesselEmail['state']['docMainFolder'];

  const stack: { id: string; name: string }[] = [...rootStack];
  for (const segment of relativeSegments) {
    const cleanSegment = cleanName(segment);
    if (!cleanSegment) continue;
    const normalized = cleanSegment.toLowerCase();
    if (stack.some(existing => existing.name.trim().toLowerCase() === normalized)) {
      continue;
    }
    stack.push({ id: `${segment}-${stack.length}`, name: cleanSegment });
  }

  if (stack.length === rootStack.length) {
    stack.push({ id: String(vessel.id || vesselName), name: vesselName });
  }

  return {
    docMainFolder,
    folderPathStack: stack,
    vesselFilter: vesselName,
  };
}

// ── Dismiss Confirm Dialog ────────────────────────────────────────────────────
// Shown when user clicks "Dismiss" on any anomaly item.
// Gives two options: "Keep it here" (close dialog) or "Move to Recycle Bin".
export function renderDismissConfirmDialog(host: VesselEmail): React.ReactElement | null {
  const anomaly = host.state.spoAnomalyDismissConfirm;
  if (!anomaly) return null;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Dismiss Confirmation"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: isMobile ? 10 : 16,
      }}
      onClick={e => { if (e.target === e.currentTarget) host.setState({ spoAnomalyDismissConfirm: null }); }}
    >
      <div style={{
        background: '#fff', borderRadius: 16, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
        padding: isMobile ? 16 : 32, width: '100%', maxWidth: 440, maxHeight: '90vh', overflowY: 'auto', position: 'relative',
      }}>
        {/* Close */}
        <button
          onClick={() => host.setState({ spoAnomalyDismissConfirm: null })}
          style={{ position: 'absolute', top: 10, right: 10, background: 'none', border: 'none', minHeight: 44, minWidth: 44, fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Cancel"
        >✕</button>

        {/* Icon + title */}
        <div style={{ textAlign: 'center', marginBottom: 20 }}>
          <div style={{ fontSize: 44, marginBottom: 10 }}>🗑️</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 17, fontWeight: 700, color: '#0f172a' }}>
            What would you like to do?
          </h3>
          <p style={{ margin: 0, fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>
            The {anomaly.item_type} <strong>"{anomaly.name}"</strong> was found outside the expected SharePoint structure.
          </p>
        </div>

        {/* Path */}
        <div style={{
          background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8,
          padding: '8px 12px', marginBottom: 22, fontSize: 11, color: '#475569',
          fontFamily: 'monospace', wordBreak: 'break-all',
        }}>
          📂 {anomaly.spo_path}
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {/* Keep it here */}
          <button
            onClick={() => host.setState({ spoAnomalyDismissConfirm: null })}
            style={{
              background: '#f8fafc', color: '#334155',
              border: '2px solid #e2e8f0', borderRadius: 10, padding: '14px 18px', minHeight: 44,
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
              transition: 'border-color 0.15s',
            }}
          >
            <span style={{ fontSize: 26 }}>📌</span>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2, color: '#0f172a' }}>Keep it here</div>
              <div style={{ fontSize: 12, color: '#64748b' }}>
                Leave this item listed in the warning section for now. You can classify or dismiss it later.
              </div>
            </div>
          </button>

          {/* Move to Recycle Bin */}
          <button
            onClick={() => host._moveAnomalyToRecycleBin(anomaly)}
            style={{
              background: 'linear-gradient(135deg, #ef4444, #dc2626)',
              color: '#fff', border: 'none', borderRadius: 10, padding: '14px 18px', minHeight: 44,
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
              transition: 'opacity 0.15s',
            }}
          >
            <span style={{ fontSize: 26 }}>🗑️</span>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2 }}>Move to Recycle Bin</div>
              <div style={{ fontSize: 12, opacity: 0.9 }}>
                Dismiss this warning and move the item record to the Recycle Bin. You can restore it from there.
              </div>
            </div>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Vessel Field Save Confirm Dialog ────────────────────────────────────────
// Shown by a vessel card's inline IMO/Hull No./Shipyard/Type field
// (renderVesselsPage/renderEditableField) on blur/select, before the value
// is written. The field can only ever be set once from the card — once
// saved it renders as locked, read-only text — so this is the one chance to
// catch a typo before it's committed.
export function renderVesselFieldConfirmDialog(host: VesselEmail): React.ReactElement | null {
  const pending = host.state.vesselFieldConfirm;
  if (!pending) return null;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Confirm Save"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: isMobile ? 10 : 16,
      }}
      onClick={e => { if (e.target === e.currentTarget) host._cancelSaveVesselField(); }}
    >
      <div style={{
        background: '#fff', borderRadius: 16, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
        padding: isMobile ? 16 : 32, width: '100%', maxWidth: 420, maxHeight: '90vh', overflowY: 'auto', position: 'relative',
      }}>
        {/* Close */}
        <button
          onClick={() => host._cancelSaveVesselField()}
          style={{ position: 'absolute', top: 10, right: 10, background: 'none', border: 'none', minHeight: 44, minWidth: 44, fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Cancel"
        >✕</button>

        {/* Icon + title */}
        <div style={{ textAlign: 'center', marginBottom: 20 }}>
          <div style={{ fontSize: 44, marginBottom: 10 }}>💾</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 17, fontWeight: 700, color: '#0f172a' }}>
            Save {pending.fieldLabel}?
          </h3>
          <p style={{ margin: 0, fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>
            This sets <strong>{pending.fieldLabel}</strong> for <strong>{pending.vesselName}</strong> to the value below.
            Once saved, this field can't be edited again from this card.
          </p>
        </div>

        {/* Value */}
        <div style={{
          background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8,
          padding: '8px 12px', marginBottom: 22, fontSize: 14, fontWeight: 700, color: '#0f172a',
          textAlign: 'center', wordBreak: 'break-word',
        }}>
          {pending.value}
        </div>

        <div style={{ display: 'flex', gap: 10 }}>
          <button
            onClick={() => host._cancelSaveVesselField()}
            style={{
              flex: 1, background: '#f8fafc', color: '#334155',
              border: '2px solid #e2e8f0', borderRadius: 10, padding: '12px 18px', minHeight: 44,
              cursor: 'pointer', fontWeight: 700, fontSize: 13,
            }}
          >
            Cancel
          </button>
          <button
            onClick={() => host._confirmSaveVesselField()}
            style={{
              flex: 1, background: 'linear-gradient(135deg, #22c55e, #16a34a)',
              color: '#fff', border: 'none', borderRadius: 10, padding: '12px 18px', minHeight: 44,
              cursor: 'pointer', fontWeight: 700, fontSize: 13,
            }}
          >
            Confirm & Save
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Classify Dialog ──────────────────────────────────────────────────────────
export function renderClassifyDialog(host: VesselEmail): React.ReactElement | null {
  const dlg = host.state.spoClassifyDialog;
  if (!dlg) return null;

  return <ClassifyModalContent host={host} dlg={dlg} />;
}

function ClassifyModalContent({ host, dlg }: { host: VesselEmail; dlg: any }): React.ReactElement {

  const { anomaly, provisioning, done, doneNormal, alreadyExisted, error } = dlg;
  const [documentScope, setDocumentScope] = React.useState<'vessels' | 'common' | 'kaizen'>('vessels');
  const [selectedVessel, setSelectedVessel] = React.useState('');
  const [selectedMainFolder, setSelectedMainFolder] = React.useState(MAIN_FOLDERS[0]?.name || '');
  const [selectedFolderPath, setSelectedFolderPath] = React.useState<string[]>([]);

  const selectedMain = MAIN_FOLDERS.find(main => main.name === selectedMainFolder);
  const categoryNodes = documentScope === 'kaizen'
    ? KAIZEN_KNOWLEDGE_BANK_TREE
    : (selectedMain ? (documentScope === 'common' ? selectedMain.commonTree : selectedMain.perVesselTree) : []);
  const categoryLabel = documentScope === 'kaizen' ? 'Kaizen category' : 'Category';
  const folderLevels: Array<{ level: number; options: typeof categoryNodes }> = [];
  let levelNodes = categoryNodes;
  for (let level = 0; level < 8 && levelNodes.length > 0; level += 1) {
    folderLevels.push({ level, options: levelNodes });
    const selected = selectedFolderPath[level];
    const selectedNode = levelNodes.find(folder => folder.name === selected);
    if (!selectedNode?.children?.length) break;
    levelNodes = selectedNode.children;
  }
  const selectedDestinationChildren = selectedFolderPath.reduce<typeof categoryNodes>((nodes, name) => {
    return nodes.find(folder => folder.name === name)?.children || [];
  }, categoryNodes);
  const canPlaceInCategory = !!(
    selectedFolderPath.length > 0 &&
    selectedDestinationChildren.length === 0 &&
    (documentScope !== 'vessels' || selectedVessel)
  );

  React.useEffect(() => {
    setDocumentScope('vessels');
    setSelectedVessel(host.state.vessels[0]?.name || '');
    setSelectedMainFolder(MAIN_FOLDERS[0]?.name || '');
    setSelectedFolderPath([]);
  }, [anomaly.id]);

  React.useEffect(() => {
    const nextNodes = documentScope === 'kaizen'
      ? KAIZEN_KNOWLEDGE_BANK_TREE
      : (selectedMain ? (documentScope === 'common' ? selectedMain.commonTree : selectedMain.perVesselTree) : []);
    const nextCategory = nextNodes[0]?.name || '';
    setSelectedFolderPath(nextCategory ? [nextCategory] : []);
  }, [documentScope, selectedMainFolder]);
  const [elapsed, setElapsed] = React.useState(0);
  const [autoClose, setAutoClose] = React.useState(8);
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  React.useEffect(() => {
    if (!provisioning) {
      setElapsed(0);
      return;
    }
    const timer = setInterval(() => setElapsed(e => e + 1), 1000);
    return () => clearInterval(timer);
  }, [provisioning]);

  React.useEffect(() => {
    if (!done && !doneNormal) { setAutoClose(8); return; }
    const timer = setInterval(() => setAutoClose(s => {
      if (s <= 1) { clearInterval(timer); handleClose(); return 0; }
      return s - 1;
    }), 1000);
    return () => clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [done, doneNormal]);
  const formatTimer = (sec: number): string => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${m < 10 ? '0' + m : m}:${s < 10 ? '0' + s : s}`;
  };

  const handleVessel = async (): Promise<void> => {
    host.setState({ spoClassifyDialog: { ...dlg, provisioning: true, error: null } });
    try {
      // 1. Create vessel record via API
      const base = host._base();
      const newVessel = await host._fetchJson(`${base}/api/vessels`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: anomaly.name, imo: null, shipyard: 'Auto-Discovered', vessel_type: 'Bulk Carrier' }),
      });
      if (!newVessel?.id) {
        // Without a real database ID we can't provision folders (the
        // backend's flat-root /provision-sites endpoint is keyed on it) —
        // surface this now instead of silently continuing with a synthetic
        // client-side id that provisioning would later reject.
        throw new Error('Vessel was not created correctly — the server did not return an ID.');
      }

      const newRecord: VesselRecord = {
        id: newVessel.id,
        name: cleanName(anomaly.name),
        imo: newVessel?.imo || '—',
        status: 'Active',
        image_url: pickRandomVesselImage('Bulk Carrier'),
      };

      // 2. Add vessel directly to top of state grid so it appears at the top of the vessel list immediately
      host.setState(prev => ({
        vessels: [newRecord, ...prev.vessels.filter(v => v.name.toLowerCase() !== newRecord.name.toLowerCase())],
      }));

      // 3. Provision SPO DMS folder structure
      await host._provisionVesselFolders(anomaly.name, newRecord.id);

      // 4. Dismiss the anomaly
      host._dismissAnomaly(anomaly.id);

      // 5. Transition to success screen!
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: true, error: null } });
    } catch (e: any) {
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, error: e?.message || 'Failed to provision vessel.' } });
    }
  };

  const handleNormal = async (): Promise<void> => {
    host.setState({ spoClassifyDialog: { ...dlg, provisioning: true, error: null } });
    try {
      const result = await host._saveNormalFolder(anomaly);
      // Dismiss anomaly warning
      host._dismissAnomaly(anomaly.id);
      // Refresh normal folder list in state
      host._loadNormalFolders();
      // Show success screen
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, doneNormal: true, alreadyExisted: result?.already_existed === true, error: null } });
    } catch (e: any) {
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, doneNormal: false, error: e?.message || 'Failed to save folder.' } });
    }
  };

  const handlePlaceInCategory = async (): Promise<void> => {
    if (!canPlaceInCategory) return;
    host.setState({ spoClassifyDialog: { ...dlg, provisioning: true, error: null } });
    try {
      const result = await host._placeAnomalyInDmsCategory(anomaly, documentScope, selectedVessel, selectedMainFolder, selectedFolderPath);
      host._dismissAnomaly(anomaly.id);
      host._loadNormalFolders();
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, doneNormal: true, alreadyExisted: result?.already_existed === true, error: null } });
    } catch (e: any) {
      host.setState({ spoClassifyDialog: { ...dlg, provisioning: false, done: false, doneNormal: false, error: e?.message || 'Failed to place folder in the selected category.' } });
    }
  };

  const handleClose = (): void => {
    if (done || doneNormal) {
      host._loadData();
    }
    host.setState({ spoClassifyDialog: null });
  };

  const renderBody = (): React.ReactElement => {
    if (doneNormal) {
      return (
        <div style={{ textAlign: 'center', padding: '12px 0' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>📁</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0f172a' }}>
            {alreadyExisted ? 'Folder Already Registered' : 'Folder Added to the DMS'}
          </h3>
          <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            <strong>"{anomaly.name}"</strong> {alreadyExisted
              ? 'was already listed as a Normal Folder in the system.'
              : 'has been saved in the selected DMS location.'}
          </p>
          <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: '#0369a1', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>📋 Saved to Normal Folders</div>
            <div>• Anomaly warning dismissed</div>
            <div>• Folder linked to the DMS</div>
            <div>• Record stored in database</div>
          </div>
          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 12 }}>
            Auto-closing in {autoClose}s…
          </div>
          <button
            onClick={handleClose}
            style={{
              background: 'linear-gradient(135deg, #475569, #334155)',
              color: '#fff', border: 'none', borderRadius: 10,
              padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(71,85,105,0.3)',
            }}
          >
            Done / View Vessels
          </button>
        </div>
      );
    }

    if (provisioning) {
      return (
        <div style={{ textAlign: 'center', padding: '8px 0' }}>
          <div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: clay.accentDark }}>
            Provisioning DMS Folder Structure…
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            Creating standard DMS folder hierarchy for <strong>"{anomaly.name}"</strong> in SharePoint Online. Please wait.
          </p>

          {/* Live Timer badge */}
          <div style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            background: clay.accentSoft, color: clay.accentDark, border: `1px solid ${clay.accentSoft}`,
            borderRadius: 20, padding: '8px 18px', fontSize: 14, fontWeight: 700, marginBottom: 20,
          }}>
            <span style={{ fontSize: 16 }}>⏱️</span>
            <span>Elapsed Time: {formatTimer(elapsed)}</span>
          </div>

          {/* Animated Progress Bar */}
          <div style={{ background: '#e2e8f0', borderRadius: 10, height: 8, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{
              background: clay.accentGradient,
              height: '100%', width: `${Math.min(96, 15 + elapsed * 12)}%`,
              transition: 'width 0.8s ease-out', borderRadius: 10,
            }} />
          </div>

    {/* Step Progress Checklist */}
          <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 }}>
            <div style={{ color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>✓</span><span>Vessel record confirmed in database</span>
            </div>
            <div style={{ color: clay.accentDark, fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>⏳</span><span>Creating SharePoint DMS folder tree (Technical &amp; Crewing, Month End, Certificates)…</span>
            </div>
            <div style={{ color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>○</span><span>Linking category subfolders &amp; permissions</span>
            </div>
          </div>

          {/* Live folder-creation feed — one line per folder as Graph confirms it */}
          {!!(host.state.folderCreationResults && host.state.folderCreationResults.length) && (
            <div style={{
              background: '#0f172a', borderRadius: 10, padding: '10px 12px', marginTop: 12,
              textAlign: 'left', fontSize: 11, fontFamily: 'monospace', maxHeight: 140,
              overflowY: 'auto', color: '#cbd5e1',
            }}>
              {host.state.folderCreationResults.map((r, i) => (
                <div key={`${r.path}-${i}`} style={{ display: 'flex', gap: 6, padding: '2px 0' }}>
                  <span>
                    {r.status === 'failed' ? '✗' : r.status === 'existed' ? '↩' : '✓'}
                  </span>
                  <span style={{
                    color: r.status === 'failed' ? '#f87171' : r.status === 'existed' ? '#94a3b8' : '#4ade80',
                    whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                  }}>
                    {r.path}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }

     if (done) {
      const failedCount = (host.state.folderCreationResults || []).filter(r => r.status === 'failed').length;
      return (
        <div style={{ textAlign: 'center', padding: '12px 0' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>{failedCount ? '⚠️' : '🎉'}</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: failedCount ? '#b45309' : '#059669' }}>
            {failedCount ? 'Vessel Provisioned With Some Issues' : 'Vessel Successfully Provisioned & Classified!'}
          </h3>
          <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            <strong>"{anomaly.name}"</strong> has been registered in the DMS database{failedCount ? ', but ' + failedCount + ' folder' + (failedCount === 1 ? '' : 's') + ' could not be created — see the log above.' : ' and its full SharePoint DMS folder tree has been created.'}
          </p>

          <div style={{ background: failedCount ? '#fffbeb' : '#ecfdf5', border: `1px solid ${failedCount ? '#fde68a' : '#a7f3d0'}`, borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: failedCount ? '#92400e' : '#065f46', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>{failedCount ? '⚠️ Provisioning Finished' : '✅ Provisioning Complete'}</div>
            <div>• Registered vessel card added to main grid</div>
            <div>• SharePoint DMS department subfolders created{failedCount ? ` (${failedCount} failed)` : ''}</div>
            <div>• Unrecognised warning dismissed</div>
          </div>

          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 12 }}>
            Auto-closing in {autoClose}s…
          </div>

          <button
            onClick={handleClose}
            style={{
              background: 'linear-gradient(135deg, #10b981, #059669)',
              color: '#fff', border: 'none', borderRadius: 10,
              padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
            }}
          >
            Done / View Vessels
          </button>
        </div>
      );
    }

    return (
      <>
        <div style={{ fontSize: 42, textAlign: 'center', marginBottom: 12 }}>
          {anomaly.item_type === 'folder' ? '📁' : '📄'}
        </div>

        <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a', textAlign: 'center' }}>
          Unclassified SharePoint {anomaly.item_type === 'folder' ? 'Folder' : 'File'}
        </h3>
        <p style={{ margin: '0 0 20px', fontSize: 13, color: '#64748b', textAlign: 'center' }}>
          We found an {anomaly.item_type === 'folder' ? 'folder' : 'file'} in SharePoint that is not currently linked to the DMS. Please choose how you would like to organize it.
        </p>

        {/* Path */}
        <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 14px', marginBottom: 20, fontSize: 11, color: '#475569', fontFamily: 'monospace', wordBreak: 'break-all' }}>
          <strong>Current location:</strong><br />📂 {anomaly.spo_path}
        </div>

        {error && (
          <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' }}>
            <span>⚠️</span><span>{error}</span>
          </div>
        )}

        {anomaly.item_type === 'folder' ? (
          <>
            <p style={{ margin: '0 0 16px', fontSize: 13, fontWeight: 600, color: '#0f172a', textAlign: 'center' }}>Choose how to organize this folder</p>
            <div style={{ display: 'flex', gap: 12, flexDirection: 'column' }}>
              {/* Vessel option */}
              <button
                onClick={() => handleVessel()}
                style={{
                  background: clay.accentGradient,
                  color: '#fff', border: 'none', borderRadius: 12, padding: '16px 20px',
                  cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
                  transition: 'opacity 0.2s', boxShadow: clay.shadowButton,
                }}
              >
                <span style={{ fontSize: 28 }}>🚢</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>
                    This is a Vessel
                  </div>
                  <div style={{ fontSize: 12, opacity: 0.9 }}>
                    Register this as a vessel and create the required DMS folder structure for it.
                  </div>
                </div>
              </button>

              <div style={{ background: '#fff', color: '#334155', border: '2px solid #e2e8f0', borderRadius: 12, padding: '16px 20px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
                  <span style={{ fontSize: 28 }}>📁</span>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', marginBottom: 2 }}>Place Under a DMS Category</div>
                    <div style={{ fontSize: 12, color: '#64748b' }}>Choose the document area and destination where this folder should be placed.</div>
                  </div>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 8, marginTop: 14 }}>
                  <select value={documentScope} onChange={e => setDocumentScope(e.target.value as 'vessels' | 'common' | 'kaizen')} style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}>
                    <option value="vessels">Vessels</option>
                    <option value="common">Common for all vessels</option>
                    <option value="kaizen">Kaizen - Knowledge Bank</option>
                  </select>
                  {documentScope === 'vessels' ? (
                    <select value={selectedVessel} onChange={e => setSelectedVessel(e.target.value)} style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}>
                      <option value="">Select Vessel</option>
                      {host.state.vessels.map(vessel => <option key={vessel.id} value={vessel.name}>{vessel.name}</option>)}
                    </select>
                  ) : <div />}
                  {documentScope !== 'kaizen' && (
                    <select value={selectedMainFolder} onChange={e => setSelectedMainFolder(e.target.value)} style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}>
                      <option value="">Select Main Folder</option>
                      {MAIN_FOLDERS.map(main => <option key={main.name} value={main.name}>{main.name}</option>)}
                    </select>
                  )}
                  {folderLevels.map(({ level, options }) => (
                    <select
                      key={level}
                      value={selectedFolderPath[level] || ''}
                      onChange={e => {
                        const nextPath = selectedFolderPath.slice(0, level);
                        if (e.target.value) nextPath.push(e.target.value);
                        setSelectedFolderPath(nextPath);
                      }}
                      style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}
                    >
                      <option value="">Select {level === 0 ? categoryLabel : `Sub-category ${level}`}</option>
                      {options.map(folder => <option key={folder.name} value={folder.name}>{folder.name}</option>)}
                    </select>
                  ))}
                </div>
                <div style={{ marginTop: 10, fontSize: 11, color: '#64748b', fontFamily: 'monospace' }}>
                  {(documentScope === 'vessels' ? `${selectedMainFolder || '[Main Folder]'} / ${selectedVessel || '[Vessel]'}` : documentScope === 'common' ? `${selectedMainFolder || '[Main Folder]'} / Common for all ships` : 'Kaizen - Knowledge Bank')}
                  {` / ${selectedFolderPath.join(' / ') || '[Category]'}`} / {anomaly.name}
                </div>
                <button onClick={handlePlaceInCategory} disabled={!canPlaceInCategory} style={{ marginTop: 12, minHeight: 44, width: isMobile ? '100%' : 'auto', background: canPlaceInCategory ? '#eff6ff' : '#f1f5f9', color: canPlaceInCategory ? '#0369a1' : '#94a3b8', border: '1px solid #bae6fd', borderRadius: 8, padding: '9px 14px', fontWeight: 700, cursor: canPlaceInCategory ? 'pointer' : 'not-allowed' }}>Place in Category</button>
              </div>

              <button
                onClick={handleNormal}
                style={{
                  background: '#f8fafc', color: '#334155',
                  border: '2px solid #e2e8f0', borderRadius: 12, padding: '16px 20px',
                  cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
                  transition: 'opacity 0.2s',
                }}
              >
                <span style={{ fontSize: 28 }}>📁</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2, color: '#0f172a' }}>
                    Keep Current Folder Location
                  </div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>
                    Keep this folder in its current SharePoint location and add it to the DMS without changing its folder path.
                  </div>
                </div>
              </button>
            </div>
          </>
        ) : (
          /* File item — provide classify + recycle options */
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {/* Place Under DMS Category */}
            <div style={{ background: '#fff', color: '#334155', border: '2px solid #e2e8f0', borderRadius: 12, padding: '16px 20px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 14 }}>
                <span style={{ fontSize: 28 }}>📁</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', marginBottom: 2 }}>Place Under a DMS Category</div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>Move this file to a DMS document folder.</div>
                </div>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 8 }}>
                <select value={documentScope} onChange={e => setDocumentScope(e.target.value as 'vessels' | 'common' | 'kaizen')} style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}>
                  <option value="vessels">Vessels</option>
                  <option value="common">Common for all vessels</option>
                  <option value="kaizen">Kaizen - Knowledge Bank</option>
                </select>
                {documentScope === 'vessels' ? (
                  <select value={selectedVessel} onChange={e => setSelectedVessel(e.target.value)} style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}>
                    <option value="">Select Vessel</option>
                    {host.state.vessels.map(vessel => <option key={vessel.id} value={vessel.name}>{vessel.name}</option>)}
                  </select>
                ) : <div />}
                {documentScope !== 'kaizen' && (
                  <select value={selectedMainFolder} onChange={e => setSelectedMainFolder(e.target.value)} style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}>
                    <option value="">Select Main Folder</option>
                    {MAIN_FOLDERS.map(main => <option key={main.name} value={main.name}>{main.name}</option>)}
                  </select>
                )}
                {folderLevels.map(({ level, options }) => (
                  <select
                    key={level}
                    value={selectedFolderPath[level] || ''}
                    onChange={e => {
                      const nextPath = selectedFolderPath.slice(0, level);
                      if (e.target.value) nextPath.push(e.target.value);
                      setSelectedFolderPath(nextPath);
                    }}
                    style={{ padding: 9, border: '1px solid #cbd5e1', borderRadius: 7, background: '#fff' }}
                  >
                    <option value="">Select {level === 0 ? categoryLabel : `Sub-category ${level}`}</option>
                    {options.map(folder => <option key={folder.name} value={folder.name}>{folder.name}</option>)}
                  </select>
                ))}
              </div>
              <div style={{ marginTop: 10, fontSize: 11, color: '#64748b', fontFamily: 'monospace' }}>
                {(documentScope === 'vessels' ? `${selectedMainFolder || '[Main Folder]'} / ${selectedVessel || '[Vessel]'}` : documentScope === 'common' ? `${selectedMainFolder || '[Main Folder]'} / Common for all ships` : 'Kaizen - Knowledge Bank')}
                {` / ${selectedFolderPath.join(' / ') || '[Category]'}`} / {anomaly.name}
              </div>
              <button onClick={handlePlaceInCategory} disabled={!canPlaceInCategory} style={{ marginTop: 12, minHeight: 44, width: isMobile ? '100%' : 'auto', background: canPlaceInCategory ? '#eff6ff' : '#f1f5f9', color: canPlaceInCategory ? '#0369a1' : '#94a3b8', border: '1px solid #bae6fd', borderRadius: 8, padding: '9px 14px', fontWeight: 700, cursor: canPlaceInCategory ? 'pointer' : 'not-allowed' }}>Place in Category</button>
            </div>

            {/* Move to Recycle Bin */}
            <button
              onClick={() => host._moveAnomalyToRecycleBin(anomaly)}
              style={{
                background: 'linear-gradient(135deg, #ef4444, #dc2626)',
                color: '#fff', border: 'none', borderRadius: 12, padding: '14px 20px',
                cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
              }}
            >
              <span style={{ fontSize: 26 }}>🗑️</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>Move to Recycle Bin</div>
                <div style={{ fontSize: 12, opacity: 0.9 }}>Remove this unclassified file and move it to the Recycle Bin.</div>
              </div>
            </button>

            {/* Dismiss only */}
            <button
              onClick={handleNormal}
              style={{
                background: '#f8fafc', color: '#334155',
                border: '2px solid #e2e8f0', borderRadius: 12, padding: '12px 20px',
                cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
              }}
            >
              <span style={{ fontSize: 26 }}>📌</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2, color: '#0f172a' }}>Keep in Current Location</div>
                <div style={{ fontSize: 12, color: '#64748b' }}>Dismiss the warning and keep this file in its current SharePoint location.</div>
              </div>
            </button>
          </div>
        )}

      </>
    );
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Classify SharePoint Item"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: isMobile ? 10 : 16,
      }}
     onClick={e => { if (e.target === e.currentTarget) handleClose(); }}
    >
      <div style={{
        background: '#fff', borderRadius: 20, padding: isMobile ? '18px 14px' : '32px 36px',
        width: isMobile ? '95vw' : 480, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto',
        boxShadow: '0 24px 64px rgba(0,0,0,0.28)', position: 'relative',
      }}>
        <button
          onClick={handleClose}
          style={{ position: 'absolute', top: 10, right: 10, background: 'none', border: 'none', minHeight: 44, minWidth: 44, fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Close (provisioning continues in the background)"
        >✕</button>
        {renderBody()}
      </div>
    </div>
  );
}


// ── Unrecognised Folders Section ─────────────────────────────────────────────


// eslint-disable-next-line @typescript-eslint/no-unused-vars
function renderUnrecognisedFolders(host: VesselEmail, items: FolderAnomalyItem[]): React.ReactElement {
  return (
    <div style={{ marginTop: 32, background: 'linear-gradient(135deg, #fffbeb 0%, #fff9e6 100%)', border: '2px solid #f59e0b', borderRadius: 16, padding: 24 }}>
      {/* Header */}
      <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <div>
          <h3 style={{ margin: '0 0 4px', fontSize: 16, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 8 }}>
            ⚠️ Folders Created Outside DMS Records
            <span style={{ background: '#f59e0b', color: '#fff', borderRadius: 20, padding: '1px 10px', fontSize: 12, fontWeight: 700 }}>{items.length}</span>
          </h3>
          <p style={{ margin: 0, fontSize: 12, color: '#a16207' }}>
            These folders were created directly in SharePoint Online at the vessel management level but are not registered in DMS.
            Please classify each one as a Vessel or a Normal Folder.
          </p>
        </div>
      </div>

      {/* Items */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {items.map(item => (
          <div
            key={item.id}
            style={{
              background: '#fff', borderRadius: 12, border: '1px solid #fde68a',
              padding: '14px 18px', display: 'flex', alignItems: 'center',
              justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
              boxShadow: '0 1px 4px rgba(245,158,11,0.08)',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 }}>
              <span style={{ fontSize: 28 }}>📁</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14, color: '#1f1f1f', marginBottom: 2 }}>{item.name}</div>
                <div style={{ fontSize: 11, color: '#78716c', fontFamily: 'monospace' }}>{item.spo_path}</div>
                {item.detected_at && (
                  <div style={{ fontSize: 10, color: '#a8a29e', marginTop: 2 }}>
                    Detected: {new Date(item.detected_at).toLocaleString()}
                  </div>
                )}
              </div>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              {/* Classify button */}
              <button
                onClick={() => host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } })}
                style={{
                  background: 'linear-gradient(135deg, #f59e0b, #d97706)',
                  color: '#fff', border: 'none', borderRadius: 8, padding: '7px 14px',
                  fontSize: 12, fontWeight: 700, cursor: 'pointer',
                  display: 'inline-flex', alignItems: 'center', gap: 5,
                }}
              >
                🔍 Classify
              </button>

              {/* Open in SPO */}
              <a
                href={getSpoVesselFolderUrl(
                  (() => {
                    const matchedVessel = item.vessel_name
                      ? host.state.vessels.find(v => v.name.toLowerCase() === item.vessel_name!.toLowerCase())
                      : undefined;
                    return matchedVessel
                      ? getVesselSharePointSiteUrl(host, matchedVessel)
                      : (host.state.documentSites.find(s => s.site_key === host.state.activeDocumentSite)?.web_url || host.props.siteUrl || '');
                  })(),
                  item.name
                )}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
                  borderRadius: 8, padding: '7px 12px', fontSize: 12, fontWeight: 600,
                  textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 4,
                }}
              >
                ↗ Open
              </a>

              {/* Dismiss → opens confirm dialog */}
              <button
                onClick={() => host.setState({ spoAnomalyDismissConfirm: item })}
                style={{
                  background: '#f1f5f9', color: '#64748b', border: '1px solid #cbd5e1',
                  borderRadius: 8, padding: '7px 12px', fontSize: 12, cursor: 'pointer',
                }}
                title="Dismiss warning"
              >
                ✕ Dismiss
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Normal Folders Section ────────────────────────────────────────────────────
function renderNormalFoldersSection(host: VesselEmail): React.ReactElement | null {
  const folders: NormalFolderRecord[] = host.state.normalFolders || [];
  if (folders.length === 0) return null;

  return (
    <div style={{ marginTop: 36 }}>
      {/* Divider */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
        <div style={{ flex: 1, height: 1, background: 'linear-gradient(to right, #94a3b8, transparent)' }} />
        <span style={{ fontSize: 12, fontWeight: 700, color: '#475569', background: '#f1f5f9', padding: '4px 14px', borderRadius: 20, border: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
          📁 Normal Folders ({folders.length})
        </span>
        <div style={{ flex: 1, height: 1, background: 'linear-gradient(to left, #94a3b8, transparent)' }} />
      </div>

      <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 14, padding: '16px 20px' }}>
        <p style={{ margin: '0 0 14px', fontSize: 12, color: '#64748b', lineHeight: 1.5 }}>
          The following folders were found in SharePoint Online but confirmed as <strong>normal (non-vessel) folders</strong> by a user. They are listed here for reference only.
        </p>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 10 }}>
          {folders.map((f, idx) => (
            <div
              key={f.id ?? `nf-${idx}`}
              style={{
                background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '12px 14px',
                display: 'flex', alignItems: 'flex-start', gap: 10,
              }}
            >
              <span style={{ fontSize: 24, flexShrink: 0, marginTop: 2 }}>
                {f.item_type === 'folder' ? '📁' : '📄'}
              </span>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {f.name}
                </div>
                <div style={{ fontSize: 10, color: '#94a3b8', fontFamily: 'monospace', marginTop: 2, wordBreak: 'break-all' }}>
                  {f.spo_path}
                </div>
                {f.detected_at && (
                  <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 4 }}>
                    Saved: {new Date(f.detected_at).toLocaleDateString()}
                  </div>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Vessel Card Provision Dialog ──────────────────────────────────────────────
// Shown when user clicks the "📁 Provision" button on a vessel card.

function renderProvisionDialog(host: VesselEmail): React.ReactElement | null {
  if (!host.state.spoProvisionDialog) return null;
  return <ProvisionModalContent host={host} dlg={host.state.spoProvisionDialog} />;
}
/** Vertical auto-scrolling ticker: shows each folder the moment the backend reports it created. */
function FolderCreationTicker({ feed }: { feed: any[] }): React.ReactElement | null {
  const scrollRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [feed.length]);

  if (!feed.length) return null;

  return (
    <div
      ref={scrollRef}
      style={{
        marginTop: 12, background: '#0f172a', borderRadius: 12, padding: '10px 14px',
        textAlign: 'left', fontSize: 11.5, fontFamily: 'monospace', maxHeight: 130,
        overflowY: 'auto', scrollBehavior: 'smooth',
      }}
    >
      {feed.map((entry, i) => (
        <div
          key={`${entry.path}-${i}`}
          style={{
            display: 'flex', alignItems: 'center', gap: 8, padding: '3px 0',
            color: entry.status === 'failed' ? '#f87171' : entry.status === 'existed' ? '#94a3b8' : '#4ade80',
          }}
        >
          <span>{entry.status === 'failed' ? '✗' : entry.status === 'existed' ? '↩' : '✓'}</span>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {entry.path.split('/').slice(-1)[0]}
          </span>
        </div>
      ))}
    </div>
  );
}
function ProvisionModalContent({ host, dlg }: { host: VesselEmail; dlg: any }): React.ReactElement {
  const { vessel, provisioning, done, error } = dlg;
  const [elapsed, setElapsed] = React.useState(0);
  const [autoClose, setAutoClose] = React.useState(10);
  const [retrying, setRetrying] = React.useState(false);
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  const isProvisioned = (
    host.state.provisionedVesselIds?.has(vessel.id) ||
    (host.state.rows && host.state.rows.some(r => r.vesselName && r.vesselName.toLowerCase() === vessel.name.toLowerCase())) ||
    vessel.is_provisioned === true
  );

  const isBusy = (provisioning || host.state.folderProvisioningVesselId === vessel.id || host.state.folderCreationBusy) && !done;
  const isDone = done || (!isBusy && ((host.state.folderCreationResults?.length ?? 0) > 0));

  React.useEffect(() => {
    if (!isBusy) { setElapsed(0); return; }
    const timer = setInterval(() => setElapsed(e => e + 1), 1000);
    return () => clearInterval(timer);
  }, [isBusy]);

  React.useEffect(() => {
    if (!isDone) { setAutoClose(10); return; }
    const timer = setInterval(() => setAutoClose(s => {
      if (s <= 1) { clearInterval(timer); handleClose(); return 0; }
      return s - 1;
    }), 1000);
    return () => clearInterval(timer);
  }, [isDone]);

  const formatTimer = (sec: number): string => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${m < 10 ? '0' + m : m}:${s < 10 ? '0' + s : s}`;
  };

  /** Re-run provisioning for this vessel. Folder creation is now a single,
   * all-or-nothing call to the backend's flat-root /provision-sites endpoint
   * (see VesselEmail.tsx's _provisionVesselFolders) rather than the old
   * client-side multi-folder Graph builder — there's no such thing as
   * "retry just the failed folders" anymore, so retry and full re-provision
   * are the same action. */
  const handleRetryInPlace = async (): Promise<void> => {
    setRetrying(true);
    try {
      const { success } = await host._provisionVesselFolders(vessel.name, vessel.id);
      host.setState({
        spoProvisionDialog: {
          ...dlg,
          provisioning: false,
          done: true,
          error: success ? null : 'Folder provisioning failed — see error below.',
        },
      });
    } finally {
      setRetrying(false);
    }
  };

  const handleProvision = async (): Promise<void> => {
    host.setState({ spoProvisionDialog: { ...dlg, provisioning: true, done: false, error: null } });
    try {
      const { success, results } = await host._provisionVesselFolders(vessel.name, vessel.id);
      const failedCount = results.filter(r => r.status === 'failed').length;
      host.setState({
        spoProvisionDialog: {
          ...dlg,
          provisioning: false,
          done: true,
          error: success ? null : `${failedCount} folder${failedCount === 1 ? '' : 's'} could not be created — see details below.`,
        },
      });
    } catch (e: any) {
      host.setState({ spoProvisionDialog: { ...dlg, provisioning: false, done: false, error: e?.message || 'Failed to provision vessel folders.' } });
    }
  };

  const handleClose = (): void => {
    host.setState({ spoProvisionDialog: null, folderCreationResults: null });
  };

  const renderBody = (): React.ReactElement => {
    if (isBusy) {
      return (
        <div style={{ textAlign: 'center', padding: '8px 0' }}>
          <div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: clay.accentDark }}>
            Provisioning DMS Folder Structure…
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            Creating standard DMS folder hierarchy for <strong>&quot;{ vessel.name }&quot;</strong> in SharePoint Online. Please wait.
          </p>

          {/* Live Timer badge */}
          <div style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            background: clay.accentSoft, color: clay.accentDark, border: `1px solid ${clay.accentSoft}`,
            borderRadius: 20, padding: '8px 18px', fontSize: 14, fontWeight: 700, marginBottom: 20,
          }}>
            <span style={{ fontSize: 16 }}>⏱️</span>
            <span>Elapsed Time: {formatTimer(elapsed)}</span>
          </div>

          {/* Animated Progress Bar */}
          <div style={{ background: '#e2e8f0', borderRadius: 10, height: 8, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{
              background: clay.accentGradient,
              height: '100%', width: `${Math.min(96, 15 + elapsed * 12)}%`,
              transition: 'width 0.8s ease-out', borderRadius: 10,
            }} />
          </div>

          {/* Step Progress Checklist */}
          <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 }}>
            <div style={{ color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>✓</span><span>Vessel record confirmed in database</span>
            </div>
            <div style={{ color: clay.accentDark, fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>⏳</span><span>Creating SharePoint DMS folder tree (Technical &amp; Crewing, Month End, Certificates)…</span>
            </div>
         <div style={{ color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 14 }}>○</span><span>Linking category subfolders &amp; permissions</span>
            </div>
          </div>

                  {/* Live Vertical News Ticker */}
          <FolderCreationTicker feed={host.state.folderCreationResults || []} />
        </div>
      );
    }

   if (isDone) {
      const spoUrl = getSpoVesselFolderUrl(host, vessel);
      const failedFolders = (host.state.folderCreationResults || []).filter(r => r.status === 'failed');
      const hasFailures = failedFolders.length > 0;
      return (
        <div style={{ textAlign: 'center', padding: '12px 0' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>{hasFailures ? '⚠️' : '🎉'}</div>
          <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: hasFailures ? '#b45309' : '#059669' }}>
            {hasFailures ? 'Vessel Folders Provisioned (with some issues)' : 'Vessel Folders Successfully Provisioned!'}
          </h3>
          <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            {hasFailures ? (
              <><strong>&quot;{vessel.name}&quot;</strong> has been provisioned, but {failedFolders.length} folder{failedFolders.length === 1 ? '' : 's'} could not be created.</>
            ) : (
              <><strong>&quot;{vessel.name}&quot;</strong> has been fully provisioned — all SharePoint DMS folder structures have been created.</>
            )}
          </p>
          {hasFailures ? (
            <div style={{ background: '#fffbeb', border: '1px solid #fde68a', borderRadius: 10, padding: '12px 16px', marginBottom: 16, fontSize: 12, color: '#92400e', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, marginBottom: 6 }}>⚠️ Folders that failed:</div>
              <div style={{ maxHeight: 100, overflowY: 'auto' }}>
                {failedFolders.map((r, i) => (
                  <div key={`${r.path}-${i}`} title={r.error}>• {r.path.split('/').pop()}</div>
                ))}
              </div>
              {/* Live retry status badge */}
              {retrying ? (
                <div style={{
                  marginTop: 10, display: 'flex', alignItems: 'center', gap: 8,
                  background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
                  borderRadius: 8, padding: '7px 12px',
                }}>
                  <span style={{ fontSize: 15 }}>⏳</span>
                  <span style={{ fontWeight: 600 }}>
                    Retrying provisioning…
                  </span>
                </div>
              ) : (
                <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  <button
                    id="btn-retry-failed-folders"
                    onClick={handleRetryInPlace}
                    style={{
                      background: 'linear-gradient(135deg,#f59e0b,#d97706)',
                      color: '#fff', border: 'none', borderRadius: 8,
                      padding: '8px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                      display: 'inline-flex', alignItems: 'center', gap: 6,
                      boxShadow: '0 3px 8px rgba(217,119,6,0.35)',
                    }}
                  >
                    🔁 Retry Provisioning
                  </button>
                </div>
              )}
            </div>
          ) : (
            <div style={{ background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '12px 16px', marginBottom: 16, fontSize: 12, color: '#065f46', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, marginBottom: 4 }}>✅ Provisioning Complete</div>
              <div>• SharePoint DMS department subfolders created</div>
              <div>• IMO: {vessel.imo || '—'}</div>
              <div>• Vessel is ready for document uploads</div>
            </div>
          )}
          {/* Auto-close countdown */}
          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 14 }}>
            Auto-closing in {autoClose}s…
          </div>
          <div style={{ background: '#e2e8f0', borderRadius: 10, height: 4, overflow: 'hidden', marginBottom: 20 }}>
            <div style={{
              background: '#10b981', height: '100%', borderRadius: 10,
              width: `${(autoClose / 10) * 100}%`, transition: 'width 1s linear',
            }} />
          </div>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button
              onClick={handleClose}
              style={{
                background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1',
                borderRadius: 10, padding: '10px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
            >
              Close
            </button>
            {spoUrl && spoUrl !== '#' && (
              <a
                href={spoUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  background: 'linear-gradient(135deg, #10b981, #059669)',
                  color: '#fff', border: 'none', borderRadius: 10,
                  padding: '10px 22px', fontSize: 13, fontWeight: 700,
                  textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 6,
                  boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
                }}
              >
                📂 Open Folder in SharePoint ↗
              </a>
            )}
          </div>
        </div>
      );
    }

    if (!provisioning && !done && isProvisioned) {
      return (
        <div style={{ textAlign: 'center', padding: '8px 0' }}>
          <div style={{ fontSize: 48, marginBottom: 10 }}>✅</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 19, fontWeight: 700, color: '#15803d' }}>
            Folders Already Provisioned
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
            The standard SharePoint DMS folder tree for <strong>&quot;{vessel.name}&quot;</strong> is already provisioned and registered.
          </p>

          <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 10, padding: '12px 16px', marginBottom: 20, fontSize: 12, color: '#166534', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>📋 Provisioned Folder Structure:</div>
            <div>• Technical &amp; Crewing — Active</div>
            <div>• Commercial &amp; Chartering — Active</div>
            <div>• Insurance — Active</div>
          </div>

          {error && (
            <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' }}>
              <span>⚠️</span><span>{error}</span>
            </div>
          )}

          <div style={{ display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button
              onClick={handleClose}
              style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
            >
              Close
            </button>
            {host.props.siteUrl && (
              <a
                href={getSpoVesselFolderUrl(host, vessel)}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  background: clay.accentGradient, color: '#fff',
                  border: 'none', borderRadius: 8, padding: '10px 20px',
                  fontSize: 13, fontWeight: 700, textDecoration: 'none',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
              >
                Open in SharePoint ↗
              </a>
            )}
            <button
              onClick={handleProvision}
              style={{
                background: 'transparent', color: clay.accentDark, border: `1px solid ${clay.accentDark}`,
                borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
            >
              🔄 Re-Provision Folders
            </button>
          </div>
        </div>
      );
    }

    return (
      <>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{ fontSize: 44, marginBottom: 10 }}>📁</div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a' }}>
            Provision SharePoint Folders
          </h3>
          <p style={{ margin: '0 0 12px', fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>
            This will create the full DMS folder hierarchy for <strong>&quot;{ vessel.name }&quot;</strong> in SharePoint Online across all departments.
          </p>
        </div>

        <div style={{ background: clay.accentSoft, border: `1px solid ${clay.accentSoft}`, borderRadius: 10, padding: '12px 16px', marginBottom: 20, fontSize: 12, color: clay.accentDark, textAlign: 'left' }}>
          <div style={{ fontWeight: 700, marginBottom: 4 }}>📋 What will be created:</div>
          <div>• Technical &amp; Crewing — monthly sub-folders + categories</div>
          <div>• Commercial — contract and invoice folders</div>
          <div>• Insurance — certificate and policy folders</div>
        </div>

        {error && (
          <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' }}>
            <span>⚠️</span><span>{error}</span>
          </div>
        )}

        <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
          <button
            onClick={handleClose}
            style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
          >
            Cancel
          </button>
          <button
            onClick={handleProvision}
            style={{
              background: clay.accentGradient,
              color: '#fff', border: 'none', borderRadius: 8,
              padding: '10px 26px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              boxShadow: clay.shadowButton,
            }}
          >
            📁 Start Provisioning
          </button>
        </div>
      </>
    );
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Provision Vessel Folders"
      style={{
        position: 'fixed', inset: 0, zIndex: 100010,
        background: 'rgba(15,23,42,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        backdropFilter: 'blur(4px)',
        padding: isMobile ? 10 : 16,
      }}
      onClick={e => { if (e.target === e.currentTarget) handleClose(); }}
    >
      <div style={{
        background: '#fff', borderRadius: 20, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
        padding: isMobile ? 16 : 32, width: '100%', maxWidth: 480, maxHeight: '90vh', overflowY: 'auto', position: 'relative',
      }}>
        <button
          onClick={handleClose}
          style={{ position: 'absolute', top: 10, right: 10, background: 'none', border: 'none', minHeight: 44, minWidth: 44, fontSize: 18, cursor: 'pointer', color: '#94a3b8' }}
          title="Close (provisioning continues in the background)"
        >✕</button>
        {renderBody()}
      </div>
    </div>
  );
}


// ── File Alert Dialog ────────────────────────────────────────────────────────
// Shown when a file is uploaded directly to SPO under the Vessels tree.
export function renderFileAlertDialog(host: VesselEmail): React.ReactElement | null {
  const dlg = host.state.spoFileAlertDialog;
  if (!dlg) return null;
  return <FileAlertDialogContent host={host} dlg={dlg} />;
}

function FileAlertDialogContent({ host, dlg }: { host: VesselEmail; dlg: import('../types/ui').SpoFileAlertDialog }): React.ReactElement {
  const { fileName, spoPath, moving, moved, error } = dlg;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  // ── Cascading selection state ──
  const [selVessel, setSelVessel] = React.useState(dlg.vesselName || '');
  const [selMainFolder, setSelMainFolder] = React.useState('');
  const [selGroupKey, setSelGroupKey] = React.useState('');

  const vessels = host.state.vessels;
  const rows = host.state.rows;

  // All unique main-folder (group) names for the selected vessel
  const mainFolders = React.useMemo(() => {
    if (!selVessel) return [];
    const norm = selVessel.trim().toLowerCase();
    return Array.from(new Set(
      rows.filter(r => r.vesselName.trim().toLowerCase() === norm && r.canUpload)
           .map(r => r.group)
    )).sort();
  }, [selVessel, rows]);

  // Sub-folder options for the selected vessel + main folder
  const subFolderOptions = React.useMemo(() => {
    if (!selVessel || !selMainFolder) return [];
    const norm = selVessel.trim().toLowerCase();
    return Array.from(
      new Map(
        rows
          .filter(r => r.vesselName.trim().toLowerCase() === norm && r.group === selMainFolder && r.canUpload)
          .map(r => [r.groupKey, { label: r.subFolderPath, groupKey: r.groupKey, uploadFolderId: r.uploadFolderId, subFolderPath: r.subFolderPath }])
      ).values()
    );
  }, [selVessel, selMainFolder, rows]);

  // Reset downstream selections when parent changes
  const onVesselChange = (v: string): void => { setSelVessel(v); setSelMainFolder(''); setSelGroupKey(''); };
  const onMainFolderChange = (v: string): void => { setSelMainFolder(v); setSelGroupKey(''); };

  const canMove = !!(selVessel && selMainFolder && selGroupKey && !moving);

  const handleMove = async (): Promise<void> => {
    if (!canMove) return;
    const target = subFolderOptions.find(o => o.groupKey === selGroupKey);
    if (!target) return;
    host.setState({ spoFileAlertDialog: { ...dlg, moving: true, error: null } });
    try {
      const { graphClient, siteId, driveId } = host.props;
      if (!graphClient || !siteId || !driveId) throw new Error('SharePoint context not available.');

      // Resolve the live SPO folder ID for the target subfolder
      const liveFolderId = host._getLiveSharePointFolderId(target.subFolderPath);
      let targetId = liveFolderId || target.uploadFolderId;

      // If targetId is a path (contains '/'), resolve it via Graph
      if (!targetId || targetId.includes('/')) {
        const parts = target.subFolderPath.split('>').map((p: string) => p.trim()).filter(Boolean);
        const vName = parts[0] || selVessel;
        const rest = parts.slice(1).join('/');
        const MAIN_FOLDER_LIST = ['Technical & Crewing', 'Commercial & Chartering', 'Insurance'];
        const candidatePaths = [
          ...MAIN_FOLDER_LIST.map(mf => `${mf}/${vName}${rest ? `/${rest}` : ''}`),
          `Vessels/Specific Vessels/${vName}/${rest}`,
          `Vessels/${vName}/${rest}`,
        ].filter(Boolean);
        for (const tryPath of candidatePaths) {
          try {
            const enc = tryPath.split('/').map((s: string) => encodeURIComponent(s)).join('/');
            const item: any = await graphClient.api(`/sites/${siteId}/drives/${driveId}/root:/${enc}?$select=id,folder`).get();
            if (item?.id && item?.folder) { targetId = item.id; break; }
          } catch { /* try next */ }
        }
      }

      if (!targetId || targetId.includes('/')) throw new Error('Could not resolve target folder in SharePoint.');

      // Verify the source file still exists at its current location
      try {
        const srcItem: any = await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${dlg.fileId}?$select=id,name,parentReference`).get();
        console.log('[VesselDMS] Move source file:', srcItem?.name, 'parentId:', srcItem?.parentReference?.id, '→ targetId:', targetId);
      } catch (verifyErr: any) {
        throw new Error(`Source file not found in SharePoint (${verifyErr?.code || verifyErr?.message}). It may have already been moved.`);
      }

      // Move the file via Graph PATCH
      let moveResult: any;
      try {
        moveResult = await graphClient.api(`/sites/${siteId}/drives/${driveId}/items/${dlg.fileId}`)
          .patch({ parentReference: { id: targetId } });
      } catch (patchErr: any) {
        const msg = patchErr?.message || patchErr?.code || String(patchErr);
        throw new Error(`Graph move failed: ${msg}`);
      }
      if (!moveResult?.id) {
        throw new Error('Graph did not confirm the file move (no item ID returned).');
      }

      // Optimistically add the file to the target folder in UI state
      host.setState(prev => {
        const newFile = { name: fileName, size: '—', date: 'Just now', pending: false, id: dlg.fileId };
        const updated = { ...prev.uploadedFilesByFolder };
        const keys = [targetId, target.groupKey, target.uploadFolderId, target.subFolderPath.trim().toLowerCase()].filter(Boolean) as string[];
        keys.forEach(k => {
          updated[k] = [...(updated[k] || []).filter((f: any) => f.name !== fileName), newFile];
        });
        return { spoFileAlertDialog: { ...dlg, moving: false, moved: true, error: null }, uploadedFilesByFolder: updated };
      });

      // Refresh the target folder from SPO using the confirmed live folder ID
      // Pass isConfirmedGraphFolderId=true so _refreshFolderFiles uses ID lookup, not path lookup
      void host._refreshFolderFiles(targetId, target.groupKey, true, true).catch(() => undefined);
      void host._syncScheduler?.triggerNow().catch(() => undefined);
    } catch (e: any) {
      host.setState({ spoFileAlertDialog: { ...dlg, moving: false, error: e?.message || 'Move failed.' } });
    }
  };

  const handleClose = (): void => { host.setState({ spoFileAlertDialog: null }); };

  const selectStyle: React.CSSProperties = {
    width: '100%', padding: '9px 12px', borderRadius: 8,
    border: '1px solid #bfdbfe', fontSize: 12, background: '#fff',
    outline: 'none', marginBottom: 10, boxSizing: 'border-box',
  };

  return (
    <div
      role="dialog" aria-modal="true" aria-label="File Uploaded Outside DMS Structure"
      style={{ position: 'fixed', inset: 0, zIndex: 100010, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 16 }}
      onClick={e => { if (e.target === e.currentTarget && !moving) handleClose(); }}
    >
      <div style={{ background: '#fff', borderRadius: 20, boxShadow: '0 24px 64px rgba(0,0,0,0.28)', padding: isMobile ? 16 : 32, width: '100%', maxWidth: 520, maxHeight: '90vh', overflowY: 'auto', position: 'relative' }}>
        {!moving && (
          <button onClick={handleClose} style={{ position: 'absolute', top: 10, right: 10, background: 'none', border: 'none', minHeight: 44, minWidth: 44, fontSize: 18, cursor: 'pointer', color: '#94a3b8' }} title="Close">✕</button>
        )}

        {moved ? (
          <div style={{ textAlign: 'center', padding: '12px 0' }}>
            <div style={{ fontSize: 52, marginBottom: 12 }}>✅</div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' }}>File Moved Successfully!</h3>
            <p style={{ margin: '0 0 8px', fontSize: 13, color: '#475569' }}>
              <strong>&quot;{fileName}&quot;</strong> has been moved to the selected DMS folder.
            </p>
            {selGroupKey && (() => {
              const t = subFolderOptions.find(o => o.groupKey === selGroupKey);
              return t ? <p style={{ margin: '0 0 20px', fontSize: 12, color: '#64748b', fontFamily: 'monospace' }}>📂 {t.subFolderPath}</p> : null;
            })()}
            <button onClick={handleClose} style={{ background: 'linear-gradient(135deg,#10b981,#059669)', color: '#fff', border: 'none', borderRadius: 10, padding: '10px 28px', fontSize: 14, fontWeight: 700, cursor: 'pointer' }}>Done</button>
          </div>
        ) : (
          <>
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
              <div style={{ width: 48, height: 48, borderRadius: 12, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, flexShrink: 0 }}>⚠️</div>
              <div>
                <h3 style={{ margin: 0, fontSize: 17, fontWeight: 700, color: '#0f172a' }}>File Uploaded Outside DMS Structure</h3>
                <p style={{ margin: '3px 0 0', fontSize: 12, color: '#64748b' }}>Detected in SharePoint but not inside a standard DMS folder.</p>
              </div>
            </div>

            {/* File info */}
            <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: '10px 14px', marginBottom: 16 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                <span style={{ fontSize: 18 }}>📄</span>
                <span style={{ fontWeight: 700, fontSize: 13, color: '#0f172a', wordBreak: 'break-all' }}>{fileName}</span>
              </div>
              <div style={{ fontSize: 11, color: '#64748b', fontFamily: 'monospace', wordBreak: 'break-all' }}>📂 {spoPath}</div>
            </div>

            {error && (
              <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '8px 12px', marginBottom: 12, fontSize: 12, color: '#dc2626' }}>⚠️ {error}</div>
            )}

            <p style={{ margin: '0 0 12px', fontSize: 13, fontWeight: 600, color: '#0f172a' }}>What would you like to do?</p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {/* Keep as is */}
              <button
                onClick={handleClose}
                style={{ background: '#f8fafc', color: '#334155', border: '2px solid #e2e8f0', borderRadius: 10, padding: '12px 16px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left' }}
              >
                <span style={{ fontSize: 24 }}>📌</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a', marginBottom: 2 }}>Keep as is</div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>Leave the file where it is. The alert stays in the bell for reference.</div>
                </div>
              </button>

              {/* Move to DMS subfolder — cascading selects */}
              <div style={{ background: 'linear-gradient(135deg,#eff6ff,#dbeafe)', border: '2px solid #bfdbfe', borderRadius: 10, padding: '14px 16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
                  <span style={{ fontSize: 24 }}>📁</span>
                  <div style={{ fontWeight: 700, fontSize: 13, color: clay.accentDark }}>Move to DMS Subfolder</div>
                </div>

                {/* Step 1: Vessel */}
                <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: clay.accentDark, marginBottom: 4, textTransform: 'uppercase' }}>1. Vessel</label>
                <select value={selVessel} onChange={e => onVesselChange(e.target.value)} style={selectStyle}>
                  <option value="">— Select vessel —</option>
                  {vessels.map(v => <option key={v.id} value={v.name}>{v.name}</option>)}
                </select>

                {/* Step 2: Main Folder / Category */}
                <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: clay.accentDark, marginBottom: 4, textTransform: 'uppercase' }}>2. Main Folder</label>
                <select value={selMainFolder} onChange={e => onMainFolderChange(e.target.value)} disabled={!selVessel || mainFolders.length === 0} style={{ ...selectStyle, opacity: !selVessel ? 0.5 : 1 }}>
                  <option value="">— Select main folder —</option>
                  {mainFolders.map(mf => <option key={mf} value={mf}>{mf}</option>)}
                </select>

                {/* Step 3: Sub-folder */}
                <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: clay.accentDark, marginBottom: 4, textTransform: 'uppercase' }}>3. Sub-folder</label>
                <select value={selGroupKey} onChange={e => setSelGroupKey(e.target.value)} disabled={!selMainFolder || subFolderOptions.length === 0} style={{ ...selectStyle, opacity: !selMainFolder ? 0.5 : 1 }}>
                  <option value="">— Select sub-folder —</option>
                  {subFolderOptions.map(o => <option key={o.groupKey} value={o.groupKey}>{o.label}</option>)}
                </select>

                <button
                  onClick={() => void handleMove()}
                  disabled={!canMove}
                  style={{
                    width: '100%',
                    background: canMove ? clay.accentGradient : clay.accentSoft,
                    color: '#fff', border: 'none', borderRadius: 8, padding: '10px',
                    fontSize: 13, fontWeight: 700,
                    cursor: canMove ? 'pointer' : 'not-allowed',
                  }}
                >
                  {moving ? '⏳ Moving…' : '📁 Move File'}
                </button>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ── Main Page Render ─────────────────────────────────────────────────────────

function renderVesselActionPicker(host: VesselEmail): React.ReactElement {
  const action = host.state.vesselActionPicker;
  const isEdit = action === 'edit';
  const isNarrow = host.state.windowWidth < 980;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
  const selectedDeleteIds = host.state.deleteVesselIds || new Set<string>();
  const selectedDeleteVessels = host.state.vessels.filter(v => selectedDeleteIds.has(v.id));
  const selectedDeleteCount = selectedDeleteVessels.length;
  const allDeleteSelected = host.state.vessels.length > 0 && host.state.vessels.every(v => selectedDeleteIds.has(v.id));

  const toggleDeleteSelection = (vesselId: string): void => {
    host.setState(prev => {
      const next = new Set(prev.deleteVesselIds);
      if (next.has(vesselId)) next.delete(vesselId);
      else next.add(vesselId);
      return { deleteVesselIds: next, modalError: null };
    });
  };

  const selectAllForDelete = (selectAll: boolean): void => {
    host.setState({
      deleteVesselIds: selectAll ? new Set(host.state.vessels.map(v => v.id)) : new Set<string>(),
      modalError: null,
    });
  };

  const proceedToDeleteModal = (): void => {
    if (selectedDeleteVessels.length === 0) {
      host.setState({ modalError: 'Select at least one vessel to delete.' });
      return;
    }
    host.setState({
      modal: 'delete',
      vesselActionPicker: null,
      selectedVessel: selectedDeleteVessels[0],
      deleteVesselIds: new Set(selectedDeleteVessels.map(v => v.id)),
      modalMsg: null,
      modalError: null,
    });
  };

  const openSingleDeleteModal = (vessel: VesselRecord): void => {
    host.setState({
      modal: 'delete',
      vesselActionPicker: null,
      selectedVessel: vessel,
      deleteVesselIds: new Set([vessel.id]),
      modalMsg: null,
      modalError: null,
    });
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={isEdit ? 'Select vessel to edit' : 'Select vessel to delete'}
      style={{ position: 'fixed', inset: 0, zIndex: 100010, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 16 }}
      onClick={e => { if (e.target === e.currentTarget) host.setState({ vesselActionPicker: null }); }}
    >
      <div style={{ background: '#fff', borderRadius: 16, width: '100%', maxWidth: isEdit ? 560 : 1020, maxHeight: '90vh', overflow: 'hidden', boxShadow: '0 24px 64px rgba(0,0,0,0.28)' }}>
        <div style={{ padding: '20px 24px 16px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0, fontSize: 18, color: '#0f172a' }}>{isEdit ? 'Select Vessel to Edit' : 'Select Vessel to Delete'}</h3>
            <p style={{ margin: '5px 0 0', fontSize: 12, color: '#64748b' }}>
              {isEdit ? 'Choose a vessel to update its details.' : 'Choose one or more vessels to move their folders and contents to the Recycle Bin.'}
            </p>
          </div>
          <button onClick={() => host.setState({ vesselActionPicker: null })} style={{ border: 'none', background: 'none', color: '#64748b', fontSize: 20, minHeight: 44, minWidth: 44, cursor: 'pointer' }} title="Close">✕</button>
        </div>
        {isEdit ? (
          <div style={{ padding: 16, maxHeight: 'calc(84vh - 100px)', overflowY: 'auto' }}>
            {host.state.vessels.length === 0 ? (
              <div style={{ padding: 20, textAlign: 'center', color: '#64748b', fontSize: 13 }}>No vessels are available.</div>
            ) : host.state.vessels.map(vessel => (
              <button
                key={vessel.id}
                onClick={() => host._openEditVessel(vessel)}
                style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 12, padding: '12px 14px', marginBottom: 8, border: '1px solid #e2e8f0', borderRadius: 10, background: '#fff', textAlign: 'left', cursor: 'pointer' }}
              >
                <span style={{ width: 36, height: 36, borderRadius: 8, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0 }}>✏️</span>
                <span style={{ minWidth: 0, flex: 1 }}>
                  <span style={{ display: 'block', fontSize: 14, fontWeight: 700, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{vessel.name}</span>
                  <span style={{ display: 'block', marginTop: 3, fontSize: 11, color: '#64748b' }}>{vessel.vessel_type || 'Vessel'}{vessel.imo ? ` | IMO ${vessel.imo}` : ''}</span>
                </span>
                <span style={{ color: clay.accentDark, fontSize: 12, fontWeight: 700 }}>Edit →</span>
              </button>
            ))}
          </div>
        ) : (
          <div style={{ padding: 16 }}>
            <div style={{
              display: 'grid',
              gridTemplateColumns: isNarrow ? '1fr' : 'minmax(0, 1.25fr) minmax(0, 0.95fr)',
              gap: 14,
              height: isNarrow ? 'calc(84vh - 132px)' : 'calc(84vh - 132px)',
              minHeight: isNarrow ? 360 : 420,
            }}>
              <div style={{ border: '1px solid #e2e8f0', borderRadius: 12, overflow: 'hidden', display: 'flex', flexDirection: 'column', background: '#fff' }}>
                <div style={{ padding: '12px 14px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, background: 'linear-gradient(180deg,#fff,#f8fafc)' }}>
                  <span style={{ fontSize: 12, color: '#334155', fontWeight: 700 }}>Vessel List</span>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span style={{ fontSize: 12, color: '#64748b', fontWeight: 600 }}>{selectedDeleteCount} selected</span>
                    <button
                      onClick={() => selectAllForDelete(!allDeleteSelected)}
                      style={{ border: '1px solid #cbd5e1', background: '#fff', color: '#334155', borderRadius: 8, padding: '5px 9px', fontSize: 12, cursor: 'pointer' }}
                    >
                      {allDeleteSelected ? 'Clear all' : 'Select all'}
                    </button>
                  </div>
                </div>

                <div style={{ padding: 10, overflowY: 'auto', flex: 1 }}>
                  {host.state.vessels.length === 0 ? (
                    <div style={{ padding: 20, textAlign: 'center', color: '#64748b', fontSize: 13 }}>No vessels are available.</div>
                  ) : host.state.vessels.map(vessel => {
                    const checked = selectedDeleteIds.has(vessel.id);
                    return (
                      <div
                        key={vessel.id}
                        onClick={() => toggleDeleteSelection(vessel.id)}
                        style={{
                          display: 'grid',
                          gridTemplateColumns: '20px 1fr auto',
                          alignItems: 'center',
                          gap: 10,
                          padding: '10px 12px',
                          marginBottom: 8,
                          border: checked ? '1px solid #fca5a5' : '1px solid #e2e8f0',
                          borderRadius: 10,
                          background: checked ? '#fff1f2' : '#fff',
                          cursor: 'pointer',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleDeleteSelection(vessel.id)}
                          onClick={e => e.stopPropagation()}
                          aria-label={`Select ${vessel.name} for deletion`}
                          style={{ width: 16, height: 16, cursor: 'pointer', accentColor: '#dc2626' }}
                        />
                        <span style={{ minWidth: 0 }}>
                          <span style={{ display: 'block', fontSize: 14, fontWeight: 700, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{vessel.name}</span>
                          <span style={{ display: 'block', marginTop: 2, fontSize: 11, color: '#64748b' }}>{vessel.vessel_type || 'Vessel'}{vessel.imo ? ` | IMO ${vessel.imo}` : ''}</span>
                        </span>
                        <button
                          onClick={e => {
                            e.stopPropagation();
                            openSingleDeleteModal(vessel);
                          }}
                          style={{ border: '1px solid #fecaca', background: '#fff', color: '#dc2626', borderRadius: 8, padding: '5px 9px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                          title="Delete this vessel only"
                        >
                          Delete only
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div style={{ border: '1px solid #fecdd3', borderRadius: 12, overflow: 'hidden', display: 'flex', flexDirection: 'column', background: 'linear-gradient(180deg,#fff1f2,#fff)' }}>
                <div style={{ padding: '12px 14px', borderBottom: '1px solid #fecdd3', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                  <span style={{ fontSize: 12, color: '#9f1239', fontWeight: 700 }}>Selected For Deletion</span>
                  <button
                    onClick={proceedToDeleteModal}
                    disabled={selectedDeleteCount === 0}
                    style={{ background: selectedDeleteCount > 0 ? 'linear-gradient(135deg, #ef4444, #dc2626)' : '#fca5a5', color: '#fff', border: 'none', borderRadius: 8, padding: '7px 11px', fontSize: 12, fontWeight: 700, cursor: selectedDeleteCount > 0 ? 'pointer' : 'not-allowed' }}
                  >
                    Delete Selected ({selectedDeleteCount})
                  </button>
                </div>

                <div style={{ padding: 12, overflowY: 'auto', flex: 1 }}>
                  {selectedDeleteCount === 0 ? (
                    <div style={{ marginTop: 8, background: '#fff', border: '1px dashed #fda4af', borderRadius: 10, padding: '12px 12px', color: '#9f1239', fontSize: 12, lineHeight: 1.45 }}>
                      Select one or more vessels from the left list to enable bulk delete.
                    </div>
                  ) : (
                    selectedDeleteVessels.map(vessel => (
                      <div key={vessel.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, padding: '8px 10px', marginBottom: 8, border: '1px solid #fecdd3', borderRadius: 9, background: '#fff' }}>
                        <span style={{ minWidth: 0 }}>
                          <span style={{ display: 'block', fontSize: 13, fontWeight: 700, color: '#881337', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{vessel.name}</span>
                          <span style={{ display: 'block', marginTop: 2, fontSize: 11, color: '#9f1239' }}>{vessel.vessel_type || 'Vessel'}{vessel.imo ? ` | IMO ${vessel.imo}` : ''}</span>
                        </span>
                        <button
                          onClick={() => toggleDeleteSelection(vessel.id)}
                          style={{ border: 'none', background: 'transparent', color: '#e11d48', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
                          title="Remove from selected"
                        >
                          Remove
                        </button>
                      </div>
                    ))
                  )}
                </div>

                {!!host.state.modalError && (
                  <div style={{ margin: '0 12px 12px', background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12 }}>{host.state.modalError}</div>
                )}

                <div style={{ padding: '0 12px 12px', display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                  <button
                    onClick={() => host.setState({ vesselActionPicker: null, modalError: null })}
                    style={{ background: '#f8fafc', border: '1px solid #cbd5e1', color: '#334155', borderRadius: 8, padding: '7px 12px', fontSize: 12, cursor: 'pointer' }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export function renderVesselsPage(host: VesselEmail): React.ReactElement {
    const {
      vessels, vesselsSearch, vesselStatusFilter, vesselTypeFilter,
      modal, selectedVessel, folderProvisioningVesselId, folderCreationError,
      folderCreationResults, panelLoading, loading,
    } = host.state;

    const normalizeSiteId = (value: string | null | undefined): string => {
      return String(value || '').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
    };

    // Resolve a raw site value to the real, CONFIGURED site_key it identifies
    // — using only host.state.documentSites (populated from GET /api/sites,
    // i.e. the backend's own env/config-driven site registry), never a
    // hardcoded list of site-name synonyms. Mirrors the backend's
    // _resolve_configured_site_key in config.py so both sides agree on what
    // counts as "the same site". Returns null if it doesn't match any
    // currently configured site.
    const resolveConfiguredSiteKey = (normalized: string): string | null => {
      if (!normalized) return null;
      const sites = host.state.documentSites || [];
      // 1. Exact match on a configured site_key.
      for (const site of sites) {
        const key = normalizeSiteId(site.site_key);
        if (key && key === normalized) return key;
      }
      // 2. The site_key appears as a whole word (handles values like
      // "Vessel DMS (dev)" -> normalized "vessel dms dev", which contains
      // the configured site_key "dev" as a token).
      const tokens = normalized.split(' ');
      for (const site of sites) {
        const key = normalizeSiteId(site.site_key);
        if (key && tokens.indexOf(key) >= 0) return key;
      }
      // 3. Match against that site's own configured display name / URL
      // (handles a value recorded under its human label rather than its
      // site_key, e.g. sp_site_name "Communication Site" or "NKSDocMan").
      for (const site of sites) {
        const key = normalizeSiteId(site.site_key);
        const candidates = [normalizeSiteId(site.sp_site_name), normalizeSiteId(site.web_url)];
        for (const candidate of candidates) {
          if (candidate && (candidate === normalized || candidate.indexOf(normalized) >= 0 || normalized.indexOf(candidate) >= 0)) {
            return key;
          }
        }
      }
      return null;
    };

    // Treat two site references as the same site only when they resolve to
    // the same CONFIGURED site (see resolveConfiguredSiteKey) — driven
    // entirely by the live /api/sites registry, never by a hardcoded list of
    // site-name synonyms. A value that happens to be "local" is only the
    // same site as "nksdocman" if the site registry itself says so;
    // otherwise they are two distinct configured sites and must not be
    // merged, however similar their names look (this is what previously
    // made every legacy "local"-tagged DMS vessel show up under NKSDocMan).
    const siteAliasMatches = (left: string | null | undefined, right: string | null | undefined): boolean => {
      const a = normalizeSiteId(left);
      const b = normalizeSiteId(right);
      if (!a || !b) return false;
      if (a === b) return true;
      const aKey = resolveConfiguredSiteKey(a);
      const bKey = resolveConfiguredSiteKey(b);
      return !!(aKey && bKey && aKey === bKey);
    };

    const filtered = vessels.filter(v => {
      if (vesselStatusFilter !== 'all' && (v.status || 'Active') !== vesselStatusFilter) return false;
      if (vesselTypeFilter && vesselTypeFilter !== 'all' && (v.vessel_type || '') !== vesselTypeFilter) return false;
      if (host.state.vesselSiteFilter && host.state.vesselSiteFilter !== 'all') {
        const selectedSiteKey = host.state.vesselSiteFilter;
        // Rows fetched with /api/vessels?site_key=<site> were already scoped
        // by the backend against the full configured site registry (e.g. a
        // vessel recorded under another site key that shares the same
        // document library) — trust that instead of re-matching them against
        // the client's partial site list. See _filterDeletedVessels.
        const fetchedForSite = (v as any)._fetched_for_site as string | undefined;
        if (fetchedForSite && fetchedForSite !== 'all') {
          if (normalizeSiteId(fetchedForSite) !== normalizeSiteId(selectedSiteKey)) return false;
        } else {
          // Optimistic / not-yet-refetched rows: match by site metadata.
          const siteMatches = (v.provisioned_site_ids || []).some(id => siteAliasMatches(String(id), selectedSiteKey)) ||
            !!(v.provisioned_site_key && siteAliasMatches(String(v.provisioned_site_key), selectedSiteKey));
          if (!siteMatches) {
            const hasNoSiteMetadata = !(v.provisioned_site_ids || []).length && !(v.provisioned_site_key || '').trim();
            if (!hasNoSiteMetadata) return false;
          }
        }
      }
      if (vesselsSearch) {
        const q = vesselsSearch.toLowerCase();
        return v.name.toLowerCase().includes(q) ||
          (v.imo || '').includes(vesselsSearch) ||
          (v.vessel_type || '').toLowerCase().includes(q) ||
          (v.shipyard || '').toLowerCase().includes(q);
      }
      return true;
    });

    const filteredDocumentSites = (host.state.documentSites || []).filter(site => {
      const siteKey = String(site.site_key || site.site_id || site.sp_site_name || '').toLowerCase();
      const siteName = String(site.sp_site_name || site.site_key || site.site_id || '').toLowerCase();
      const activeKey = String(host.state.activeDocumentSite || '').toLowerCase();
      const isDefaultSite = siteKey === activeKey || site.is_primary === true;
      const isManagedSite =
        isDefaultSite ||
        siteKey.includes('dev') ||
        siteKey.includes('communication') ||
        siteKey.includes('local') ||
        siteKey.includes('docman') ||
        siteKey.includes('nks') ||
        siteKey.includes('external') ||
        siteName.includes('dev') ||
        siteName.includes('communication') ||
        siteName.includes('local') ||
        siteName.includes('docman') ||
        siteName.includes('nks') ||
        siteName.includes('external');
      return isManagedSite;
    });

    const siteOptions = Array.from(new Map(
      filteredDocumentSites.map(site => [String(site.site_key || site.site_id || site.sp_site_name || '').toLowerCase(), site])
    ).values()).map(site => ({
      key: String(site.site_key || site.site_id || site.sp_site_name || ''),
      label: site.sp_site_name || site.site_key || site.site_id || 'SharePoint Site',
    }));

    // All unique vessel types for filter dropdown
    const allTypes = Array.from(new Set(vessels.map(v => v.vessel_type).filter(Boolean))) as string[];

    const isLoading = panelLoading || (loading && vessels.length === 0);

    // Segregated anomalies
    const allAnomalies = host.state.folderAnomalies || [];
    const vesselLevelFolders = allAnomalies.filter(a => a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'folder');
    const vesselLevelFiles   = allAnomalies.filter(a => a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'file');

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>

        {/* ── Header ── */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24, flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 28, fontWeight: 700, color: clay.text, display: 'flex', alignItems: 'center', gap: 12 }}>
              Vessels
              {vessels.length > 0 && (
                <span style={{ background: clay.accentGradient, color: '#fff', borderRadius: 20, padding: '3px 14px', fontSize: 14, fontWeight: 700, boxShadow: clay.shadowIcon }}>
                  {vessels.length}
                </span>
              )}
              {(vesselLevelFolders.length + vesselLevelFiles.length) > 0 && (
                <span
                  onClick={() => host._toggleAlertBell()}
                  style={{ background: '#fef3c7', color: '#92400e', borderRadius: 20, padding: '3px 12px', fontSize: 12, fontWeight: 700, border: '1px solid #f59e0b', display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}
                  title="SPO items detected outside DMS structure — click to view in Alerts"
                >
                  ⚠️ {vesselLevelFolders.length + vesselLevelFiles.length} SPO item{vesselLevelFolders.length + vesselLevelFiles.length !== 1 ? 's' : ''} — view in Alerts →
                </span>
              )}
            </h2>
            <p style={{ margin: '6px 0 0', fontSize: 14, color: clay.textMuted }}>
              Manage fleet vessels, provision SharePoint folders, and view documents.
            </p>
          </div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            <button
              onClick={() => host._goToView('vessels').catch(() => undefined)}
              title="Reload vessel list from database"
              aria-label="Refresh vessel list"
              onMouseEnter={e => { e.currentTarget.style.background = clay.surfaceHover; e.currentTarget.style.boxShadow = clay.shadowRaisedHover; }}
              onMouseLeave={e => { e.currentTarget.style.background = clay.surface; e.currentTarget.style.boxShadow = clay.shadowRaised; }}
              style={{ width: 42, height: 42, padding: 0, background: clay.surface, color: clay.textMuted, border: 'none', borderRadius: '50%', fontSize: 16, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', boxShadow: clay.shadowRaised, transition: 'all 0.18s ease' }}
            >
              <Icon iconName="Refresh" />
            </button>
            <button
              onClick={() => host._openVesselActionPicker('edit')}
              title="Edit vessel"
              aria-label="Edit vessel"
              onMouseEnter={e => { e.currentTarget.style.background = clay.surfaceHover; e.currentTarget.style.boxShadow = clay.shadowRaisedHover; }}
              onMouseLeave={e => { e.currentTarget.style.background = clay.surface; e.currentTarget.style.boxShadow = clay.shadowRaised; }}
              style={{ width: 42, height: 42, padding: 0, background: clay.surface, color: clay.accentDark, border: 'none', borderRadius: '50%', fontSize: 16, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', boxShadow: clay.shadowRaised, transition: 'all 0.18s ease' }}
            >
              <Icon iconName="Edit" />
            </button>
            <button
              onClick={() => host._openVesselActionPicker('delete')}
              title="Delete vessel"
              aria-label="Delete vessel"
              onMouseEnter={e => { e.currentTarget.style.background = '#f6dbd5'; e.currentTarget.style.boxShadow = '10px 10px 22px rgba(168,90,66,0.22), -10px -10px 20px rgba(255,255,255,0.92)'; }}
              onMouseLeave={e => { e.currentTarget.style.background = clay.surface; e.currentTarget.style.boxShadow = clay.shadowRaised; }}
              style={{ width: 42, height: 42, padding: 0, background: clay.surface, color: '#a4262c', border: 'none', borderRadius: '50%', fontSize: 16, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', boxShadow: clay.shadowRaised, transition: 'all 0.18s ease' }}
            >
              <Icon iconName="Delete" />
            </button>
            {/* Manual "Sync Vessels from SharePoint" button removed — the
                backend already runs this reconciliation automatically in
                the background on every vessel-list fetch, throttled per
                site (see _maybeAutoSyncVesselsFromSharePoint). */}
            <button
              onClick={() => host._openCreate()}
              title="Create new vessel"
              aria-label="Create new vessel"
              onMouseEnter={e => { e.currentTarget.style.background = clay.accentGradientHover; e.currentTarget.style.boxShadow = '0 14px 26px rgba(199,122,62,0.42), inset 0 2px 3px rgba(255,255,255,0.45), inset 0 -3px 6px rgba(150,89,42,0.28)'; }}
              onMouseLeave={e => { e.currentTarget.style.background = clay.accentGradient; e.currentTarget.style.boxShadow = clay.shadowButton; }}
              style={{ width: 42, height: 42, padding: 0, background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: '50%', fontSize: 18, fontWeight: 700, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', boxShadow: clay.shadowButton, transition: 'all 0.18s ease' }}
            >
              <Icon iconName="Add" />
            </button>
          </div>
        </div>

        {/* ── Search & Filter Bar ── */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20, flexWrap: 'wrap' }}>
          <div style={{ position: 'relative', flex: '1 1 220px', minWidth: 180 }}>
            <span style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 14 }}>🔍</span>
            <input
              type="text"
              placeholder="Search by name, IMO, type, shipyard…"
              value={vesselsSearch}
              onChange={e => host.setState({ vesselsSearch: e.target.value })}
              style={{ width: '100%', padding: '10px 14px 10px 38px', borderRadius: clay.radiusButton, border: 'none', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: clay.surface, boxShadow: 'inset 2px 2px 5px rgba(120,190,185,0.22), inset -2px -2px 4px rgba(255,255,255,0.9)', color: clay.text }}
            />
          </div>
          <select
            value={vesselStatusFilter}
            onChange={e => host.setState({ vesselStatusFilter: e.target.value })}
            style={{ padding: '10px 14px', borderRadius: clay.radiusButton, border: 'none', fontSize: 13, background: clay.surface, outline: 'none', minWidth: 130, color: clay.text, boxShadow: 'inset 2px 2px 5px rgba(120,190,185,0.22), inset -2px -2px 4px rgba(255,255,255,0.9)' }}
          >
            <option value="all">All Status</option>
            <option value="Active">Active</option>
            <option value="In Maintenance">In Maintenance</option>
            <option value="Inactive">Inactive</option>
          </select>
          <select
            value={vesselTypeFilter || 'all'}
            onChange={e => host.setState({ vesselTypeFilter: e.target.value })}
            style={{ padding: '10px 14px', borderRadius: clay.radiusButton, border: 'none', fontSize: 13, background: clay.surface, outline: 'none', minWidth: 140, color: clay.text, boxShadow: 'inset 2px 2px 5px rgba(120,190,185,0.22), inset -2px -2px 4px rgba(255,255,255,0.9)' }}
          >
            <option value="all">All Types</option>
            {allTypes.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          <select
            value={host.state.vesselSiteFilter || 'all'}
            onChange={e => {
              // The vessel list fetched by _loadData is already scoped to a
              // single site_key (see _loadData Step 1) — switching sites here
              // must re-fetch /api/vessels?site_key=<new site>, not just
              // re-filter the previous site's already-loaded vessels
              // client-side (which is what this used to do, and why picking
              // NissenKaiunExternal / NKSDocMan without hitting "Sync" showed
              // nothing and sent no request).
              host.setState({ vesselSiteFilter: e.target.value }, () => { void host._loadData(true); });
            }}
            style={{ padding: '10px 14px', borderRadius: clay.radiusButton, border: 'none', fontSize: 13, background: clay.surface, outline: 'none', minWidth: 160, color: clay.text, boxShadow: 'inset 2px 2px 5px rgba(120,190,185,0.22), inset -2px -2px 4px rgba(255,255,255,0.9)' }}
          >
            <option value="all">All SharePoint Sites</option>
            {siteOptions.map(site => <option key={site.key} value={site.key}>{site.label}</option>)}
          </select>
          {(vesselsSearch || vesselStatusFilter !== 'all' || (vesselTypeFilter && vesselTypeFilter !== 'all') || (host.state.vesselSiteFilter && host.state.vesselSiteFilter !== 'all')) && (
            <button
              onClick={() => {
                const siteFilterWasScoped = !!host.state.vesselSiteFilter && host.state.vesselSiteFilter !== 'all';
                host.setState(
                  { vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all', vesselSiteFilter: 'all' },
                  () => { if (siteFilterWasScoped) void host._loadData(true); },
                );
              }}
              style={{ background: 'transparent', color: clay.textMuted, border: `1px solid ${clay.accentSoft}`, borderRadius: clay.radiusButton, padding: '9px 14px', fontSize: 13, cursor: 'pointer' }}
            >
              ✕ Clear
            </button>
          )}
          <span style={{ marginLeft: 'auto', fontSize: 12, color: clay.textMuted, fontWeight: 500 }}>
            {filtered.length} vessel{filtered.length !== 1 ? 's' : ''}
          </span>
        </div>

        {/* ── Folder creation status banners ── */}
        {folderCreationError && (
          <div style={{ marginBottom: 12, background: '#fde7e9', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#a4262c', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>⚠ {folderCreationError}</span>
            <button onClick={() => host.setState({ folderCreationError: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#a4262c', fontWeight: 700 }}>✕</button>
          </div>
        )}
        {folderCreationResults && folderCreationResults.length > 0 && !folderCreationError && (() => {
          const createdCount = folderCreationResults.filter(r => r.status === 'created').length;
          const existedCount = folderCreationResults.filter(r => r.status === 'existed').length;
          if (createdCount === 0 && existedCount === 0) return null;

          // Determine the vessel that was just provisioned for correct link generation
          const provisionedVessel = selectedVessel ||
            vessels.find(v => v.id === host.state.folderProvisioningVesselId) ||
            vessels.find(v => (v.provisioned_site_ids || []).includes(host.state.activeDocumentSite || ''));
          const bannerSiteUrl = provisionedVessel
            ? getVesselSharePointSiteUrl(host, provisionedVessel)
            : (host.state.activeDocumentSite
                ? (host.state.documentSites.find(s => s.site_key === host.state.activeDocumentSite)?.web_url || host.props.siteUrl)
                : host.props.siteUrl);
          const bannerHref = getSpoVesselFolderUrl(host, provisionedVessel, bannerSiteUrl);
          return (
            <div style={{ marginBottom: 12, background: '#dff6dd', border: '1px solid #86efac', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#107c10', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                <span>✅ SharePoint folders provisioned — {createdCount} created, {existedCount} already existed.</span>
                {bannerHref && bannerHref !== '#' && (
                  <a
                    href={bannerHref}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{ color: clay.accentDark, fontWeight: 600, textDecoration: 'underline', display: 'inline-flex', alignItems: 'center', gap: 3 }}
                  >
                    Open SharePoint Folder ↗
                  </a>
                )}
              </div>
              <button onClick={() => host.setState({ folderCreationResults: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#107c10', fontWeight: 700 }}>✕</button>
            </div>
          );
        })()}

        {/* ── Vessel sync (SharePoint auto-discovery) status banners ── */}
        {host.state.vesselSyncError && (
          <div style={{ marginBottom: 12, background: '#fde7e9', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#a4262c', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>⚠ {host.state.vesselSyncError}</span>
            <button onClick={() => host.setState({ vesselSyncError: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#a4262c', fontWeight: 700 }}>✕</button>
          </div>
        )}
        {/* Success/summary banner for the manual sync removed along with the
            button above — it only ever surfaced host.state.vesselSyncSummary,
            which nothing else sets. The error banner above still applies to
            _confirmDiscoveredVessel's own failures. */}

        {/* ── Loading State ── */}
        {isLoading ? (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: clay.surface, borderRadius: clay.radiusCard, border: 'none', boxShadow: clay.shadowRaised }}>
            <div style={{
              width: 40, height: 40, border: `3px solid ${clay.accentSoft}`,
              borderTopColor: clay.accent, borderRadius: '50%',
              animation: 'spin 0.8s linear infinite',
            }} />
            <p style={{ margin: 0, fontSize: 14, color: clay.textMuted, fontWeight: 500 }}>Loading vessels from database…</p>
            <p style={{ margin: 0, fontSize: 12, color: clay.textMuted }}>Fetching vessel records and folder structure</p>
          </div>
        ) : filtered.length === 0 ? (
          /* ── Empty State ── */
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: clay.surface, borderRadius: clay.radiusCard, border: `2px dashed ${clay.accentSoft}` }}>
            <span style={{ fontSize: 48 }}>🚢</span>
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: 0, fontSize: 16, fontWeight: 700, color: clay.text }}>
                {vessels.length === 0 ? 'No vessels yet' : 'No vessels match your filter'}
              </p>
              <p style={{ margin: '6px 0 0', fontSize: 13, color: clay.textMuted }}>
                {vessels.length === 0
                  ? 'Create your first vessel to provision its SharePoint folder structure.'
                  : 'Try clearing the search or filters.'}
              </p>
            </div>
            {vessels.length === 0 && (
              <button
                onClick={() => host._openCreate()}
                style={{ background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: clay.radiusButton, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer', boxShadow: clay.shadowButton }}
              >
                + Create First Vessel
              </button>
            )}
          </div>
        ) : (
          /* ── Vessel Cards Grid ── */
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 20 }}>
            {filtered.map(vessel => {
              const isSelected = selectedVessel?.id === vessel.id;
              const status = vessel.status || 'Active';
              // "Found in SharePoint" = a real vessel folder discovered via the
              // site's Term Store (see GET /api/vessels), not yet added as a
              // DMS vessel record — styled like "In Maintenance" (amber, needs
              // attention) rather than the red "danger" pill everything else
              // falls into.
              const isSharePointOnly = vessel.source === 'sharepoint';
              const statusText  = status === 'Active' ? clay.pillActiveText : (status === 'In Maintenance' || isSharePointOnly) ? clay.pillWarnText : clay.pillDangerText;
              const statusBg    = status === 'Active' ? clay.pillActiveBg : (status === 'In Maintenance' || isSharePointOnly) ? clay.pillWarnBg : clay.pillDangerBg;
              const statusShadow = status === 'Active' ? clay.pillActiveShadow : (status === 'In Maintenance' || isSharePointOnly) ? clay.pillWarnShadow : clay.pillDangerShadow;
              const isProvisioning = folderProvisioningVesselId === vessel.id;
              const imgSrc = getVesselImageForId(vessel.id);

              return (
                <div
                  key={vessel.id}
                  className="vessel-card"
                  onClick={() => host.setState({ selectedVessel: isSelected ? null : vessel })}
                  style={{
                    background: clay.surface,
                    borderRadius: clay.radiusCard,
                    border: 'none',
                    boxShadow: isSelected
                      ? `0 0 0 3px ${clay.accent}, ${clay.shadowRaisedHover}`
                      : clay.shadowRaised,
                    overflow: 'hidden',
                    cursor: 'pointer',
                    transition: 'all 0.18s ease',
                    display: 'flex',
                    flexDirection: 'column',
                  }}
                  onMouseEnter={e => { if (!isSelected) { e.currentTarget.style.boxShadow = clay.shadowRaisedHover; e.currentTarget.style.transform = 'translateY(-1px)'; } }}
                  onMouseLeave={e => { if (!isSelected) { e.currentTarget.style.boxShadow = clay.shadowRaised; e.currentTarget.style.transform = 'translateY(0)'; } }}
                >
                  {/* Card image header — professional tone-on-tone banner
                      generated in vesselImagePool.ts (deep gradient + faint
                      engraved texture + a minimal type glyph watermark), no
                      cartoon illustration overlay. */}
                  <div style={{ position: 'relative', height: 140, overflow: 'hidden', background: '#101826' }}>
                    <img
                      className="vessel-card-image"
                      src={resolveImgUrl(imgSrc)}
                      alt={vessel.name}
                      style={{ width: '100%', height: '100%', objectFit: 'cover', display: 'block' }}
                    />
                    <div style={{
                      position: 'absolute', inset: 0,
                      background: 'linear-gradient(to bottom, rgba(8,14,24,0.05) 0%, rgba(8,14,24,0.15) 55%, rgba(8,14,24,0.78) 100%)',
                    }} />
                    {/* Status badge */}
                    <span style={{
                      position: 'absolute', top: 12, right: 12,
                      background: statusBg, color: statusText,
                      borderRadius: 20, padding: '3px 12px', fontSize: 12, fontWeight: 700,
                      letterSpacing: '0.2px',
                      border: 'none', boxShadow: statusShadow,
                    }}>
                      {status}
                    </span>
                    {/* Selection check */}
                    {isSelected && (
                      <span style={{
                        position: 'absolute', top: 12, left: 12,
                        background: clay.accentGradient, color: '#fff',
                        borderRadius: '50%', width: 26, height: 26,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        fontSize: 14, fontWeight: 700, boxShadow: clay.shadowIcon,
                      }}>✓</span>
                    )}
                    {/* Vessel name + type, set on a legibility scrim */}
                    <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, padding: '20px 16px 12px' }}>
                      <p style={{
                        margin: 0, fontSize: 17, fontWeight: 700, color: '#fff', letterSpacing: '0.1px',
                        textShadow: '0 1px 6px rgba(0,0,0,0.5)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                      }}>
                        {vessel.name}
                      </p>
                      {(vessel.vessel_type || vessel.imo) && (
                        <p style={{
                          margin: '3px 0 0', fontSize: 11.5, fontWeight: 600, color: 'rgba(255,255,255,0.78)',
                          letterSpacing: '0.4px', textTransform: 'uppercase',
                          textShadow: '0 1px 4px rgba(0,0,0,0.5)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                        }}>
                          {[vessel.vessel_type, vessel.imo ? `IMO ${vessel.imo}` : null].filter(Boolean).join(' · ')}
                        </p>
                      )}
                    </div>
                  </div>

                  {/* Card body */}
                  <div style={{ padding: '16px 18px', flex: 1, display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {(() => {
                      const draftKey = host._discoveredVesselKey(vessel);
                      const draft = host.state.discoveredVesselDrafts[draftKey];
                      const imoVal = isSharePointOnly ? (draft?.imo ?? vessel.imo ?? '') : (vessel.imo || '—');
                      const hullVal = isSharePointOnly ? (draft?.hull_number ?? vessel.hull_number ?? '') : (vessel.hull_number || '—');
                      const editableInputStyle: React.CSSProperties = {
                        width: '100%', marginTop: 2, padding: '4px 6px', fontSize: 13, fontWeight: 600, color: clay.text,
                        border: `1px solid ${clay.accentSoft}`, borderRadius: 6, boxSizing: 'border-box', background: '#fff',
                      };

                      // Regular ("dms"-source) vessel cards: any of these four
                      // fields that's still empty gets a directly-editable
                      // control here. Blurring (or, for Type, selecting) it
                      // opens a confirmation dialog (host._requestSaveVesselField
                      // / renderVesselFieldConfirmDialog) rather than saving
                      // straight away, and once the PATCH /api/vessels/{id}
                      // (host._saveVesselField) succeeds the field has a saved
                      // value and is rendered locked/read-only below — each of
                      // these fields can be set exactly once from the card, not
                      // edited back and forth. This is different from the
                      // deferred discoveredVesselDrafts/"Confirm" flow the
                      // "Found in SharePoint" cards use above.
                      const fieldDraft = host.state.vesselFieldDrafts[vessel.id] || {};
                      const renderEditableField = (
                        field: 'imo' | 'hull_number' | 'shipyard' | 'vessel_type',
                        placeholder: string,
                      ): React.ReactElement => {
                        const fieldKey = `${vessel.id}:${field}`;
                        const saving = !!host.state.vesselFieldSaving[fieldKey];
                        const error = host.state.vesselFieldError[fieldKey];
                        const savedVal = (vessel as any)[field] || '';

                        // Already saved once — lock it. No more edits from the
                        // card; a correction has to go through the full "Edit
                        // Vessel" form instead of this one-shot inline field.
                        if (savedVal) {
                          return (
                            <p style={{
                              margin: '2px 0 0', fontSize: 13, fontWeight: 600, color: clay.text,
                              whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                            }}>
                              {savedVal}
                            </p>
                          );
                        }

                        // Fall back to '' when there's no in-progress edit yet
                        // (the field is empty, or the field is not empty but
                        // isn't rendered as an input at all — see above).
                        const draftVal = fieldDraft[field] ?? '';
                        return (
                          <>
                            {field === 'vessel_type' ? (
                              <select
                                value={draftVal}
                                disabled={saving}
                                onClick={e => e.stopPropagation()}
                                onChange={e => {
                                  const val = e.target.value;
                                  host._updateVesselFieldDraft(vessel, field, val);
                                  if (val) host._requestSaveVesselField(vessel, field, val);
                                }}
                                style={editableInputStyle}
                              >
                                <option value="">{placeholder}</option>
                                {VESSEL_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                              </select>
                            ) : (
                              <input
                                type="text"
                                value={draftVal}
                                placeholder={placeholder}
                                disabled={saving}
                                onClick={e => e.stopPropagation()}
                                onChange={e => host._updateVesselFieldDraft(vessel, field, e.target.value)}
                                onBlur={e => host._requestSaveVesselField(vessel, field, e.target.value)}
                                style={editableInputStyle}
                              />
                            )}
                            {error && (
                              <p style={{ margin: '2px 0 0', fontSize: 10, color: '#c0392b' }}>{error}</p>
                            )}
                          </>
                        );
                      };

                      return (
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px 16px' }}>
                          <div>
                            <span style={{ fontSize: 10, color: clay.textMuted, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>IMO</span>
                            {isSharePointOnly ? (
                              <input
                                type="text"
                                value={imoVal}
                                placeholder="7-digit IMO"
                                onClick={e => e.stopPropagation()}
                                onChange={e => host._updateDiscoveredVesselDraft(vessel, 'imo', e.target.value)}
                                style={editableInputStyle}
                              />
                            ) : (
                              // Editable only while empty, and only once — see
                              // renderEditableField above.
                              renderEditableField('imo', '7-digit IMO')
                            )}
                          </div>
                          <div>
                            <span style={{ fontSize: 10, color: clay.textMuted, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Type</span>
                            {!isSharePointOnly ? (
                              renderEditableField('vessel_type', 'Select type')
                            ) : (
                              <p style={{ margin: '2px 0 0', fontSize: 13, color: clay.text, fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{vessel.vessel_type || '—'}</p>
                            )}
                          </div>
                          <div>
                            <span style={{ fontSize: 10, color: clay.textMuted, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Shipyard</span>
                            {!isSharePointOnly ? (
                              renderEditableField('shipyard', 'Shipyard')
                            ) : (
                              <p style={{ margin: '2px 0 0', fontSize: 13, color: clay.text, fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{vessel.shipyard || '—'}</p>
                            )}
                          </div>
                          <div>
                            <span style={{ fontSize: 10, color: clay.textMuted, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Hull No.</span>
                            {isSharePointOnly ? (
                              <input
                                type="text"
                                value={hullVal}
                                placeholder="Hull number"
                                onClick={e => e.stopPropagation()}
                                onChange={e => host._updateDiscoveredVesselDraft(vessel, 'hull_number', e.target.value)}
                                style={editableInputStyle}
                              />
                            ) : (
                              renderEditableField('hull_number', 'Hull number')
                            )}
                          </div>
                        </div>
                      );
                    })()}

                    <div style={{ background: clay.bg, borderRadius: 12, padding: '8px 10px', display: 'flex', flexDirection: 'column', gap: 4, boxShadow: 'inset 2px 2px 5px rgba(120,190,185,0.2), inset -2px -2px 4px rgba(255,255,255,0.85)' }}>
                      <span style={{ fontSize: 10, color: clay.textMuted, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>SharePoint Site</span>
                      <span style={{ fontSize: 12, color: clay.text, fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                        {(() => {
                          const sites = host.state.documentSites || [];
                          // Prefer the site this row was actually fetched
                          // under (_fetched_for_site — set server-side by
                          // /api/vessels?site_key=..., already resolved
                          // against the full configured site registry; see
                          // _filterDeletedVessels in VesselEmail.tsx) over
                          // provisioned_site_key. Two site_key prefixes can
                          // legitimately share one physical SharePoint
                          // library (e.g. "local" and "nksdocman" pointing
                          // at the same drive_id) — a vessel recorded under
                          // the "local" alias still lives on the site the
                          // user knows as NKSDocMan, and provisioned_site_key
                          // alone isn't a presentable site name.
                          const fetchedForSite = (vessel as any)._fetched_for_site as string | undefined;
                          const match = fetchedForSite && fetchedForSite !== 'all'
                            ? sites.find(site => normalizeSiteId(site.site_key) === normalizeSiteId(fetchedForSite))
                            : sites.find(site =>
                                (vessel.provisioned_site_ids || []).some(id => siteAliasMatches(String(id), site.site_key)) ||
                                (!!vessel.provisioned_site_key && siteAliasMatches(vessel.provisioned_site_key, site.site_key))
                              );
                          return match?.sp_site_name || match?.site_key || vessel.provisioned_site_key || 'Active site';
                        })()}
                      </span>
                    </div>

                    {/* Always shown (not just when set) so a vessel whose
                        vessel_folder_path is still missing is visibly "—"
                        instead of silently dropping the row — most existing
                        gaps are backfilled server-side from the Folder table
                        at read time (see list_vessels in real_backend.py). */}
                    <div style={{ background: clay.bg, borderRadius: 12, padding: '8px 10px', display: 'flex', flexDirection: 'column', gap: 4, boxShadow: 'inset 2px 2px 5px rgba(120,190,185,0.2), inset -2px -2px 4px rgba(255,255,255,0.85)' }}>
                      <span style={{ fontSize: 10, color: clay.textMuted, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Created Path</span>
                      <span style={{ fontSize: 11, color: clay.text, fontFamily: 'monospace', wordBreak: 'break-word' }}>{vessel.vessel_folder_path || '—'}</span>
                    </div>

                    {isSharePointOnly && (() => {
                      const confirmKey = host._discoveredVesselKey(vessel);
                      const isConfirming = host.state.confirmingVesselKey === confirmKey;
                      const confirmError = host.state.discoveredVesselConfirmError[confirmKey];
                      return (
                        <>
                          <button
                            onClick={e => { e.stopPropagation(); host._confirmDiscoveredVessel(vessel).catch(() => undefined); }}
                            disabled={isConfirming}
                            title="Save the IMO/Hull No. above and register this SharePoint folder as a DMS vessel record (no new folder is created)"
                            style={{
                              marginTop: 2, padding: '8px 12px', borderRadius: clay.radiusButton, border: 'none',
                              background: clay.accentGradient, color: '#fff', fontSize: 12, fontWeight: 700,
                              cursor: isConfirming ? 'default' : 'pointer', opacity: isConfirming ? 0.7 : 1,
                              boxShadow: clay.shadowButton,
                            }}
                          >
                            {isConfirming ? 'Confirming…' : '✓ Confirm Vessel'}
                          </button>
                          {confirmError && (
                            <p style={{ margin: '2px 0 0', fontSize: 11, color: '#c0392b', fontWeight: 600 }}>
                              ⚠ {confirmError}
                            </p>
                          )}
                        </>
                      );
                    })()}

                    {/* Action buttons */}
                    <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                      <button
                        disabled={isSharePointOnly}
                        onClick={async e => {
                          e.stopPropagation();
                          // A "Found in SharePoint" row has no DB vessel_id yet and
                          // its vessel_folder_path is just its bare folder name, not
                          // a real Graph drive-item id — the placeholder-id-to-real-id
                          // resolution below is best-effort name matching and can
                          // land on the wrong folder or none at all, which is what
                          // produced "Couldn't load this folder" for these cards.
                          // Confirm the vessel first (button above) so it has a real
                          // DB record before browsing its documents.
                          if (isSharePointOnly) return;
                          const nav = resolveVesselDocumentNavigation(host, vessel);
                          const navSite = nav.folderPathStack.length >= 2
                            ? (host.state.documentSites || []).find(site =>
                                `site:${site.site_id || site.site_key}` === nav.folderPathStack[1].id ||
                                site.site_key === vessel.provisioned_site_key ||
                                (vessel.provisioned_site_ids || []).some(id => String(id).toLowerCase() === String(site.site_key || '').toLowerCase())
                              )
                            : null;
                          let folderPathStack = nav.folderPathStack;
                          let resolvedDriveId: string | null = navSite?.drive_id || null;
                          if (navSite?.site_id && nav.folderPathStack[0]?.id === 'sites_root') {
                            try {
                              const liveDrives = await host._loadSiteDrives(navSite.site_id);
                              const preferredDrive = liveDrives.find(d => /^(documents|shared documents)$/i.test(d.name.trim()));
                              const liveDrive = preferredDrive || liveDrives.find(d => d.id === navSite.drive_id) || liveDrives[0];
                              if (liveDrive) {
                                resolvedDriveId = liveDrive.id;
                                folderPathStack = nav.folderPathStack.map((node, index) => index === 2
                                  ? { id: `drive:${liveDrive.id}`, name: liveDrive.name || node.name }
                                  : node);
                                host.setState({
                                  documentSites: host.state.documentSites.map(site => site.site_key === navSite.site_key
                                    ? { ...site, drive_id: liveDrive.id, default_library_name: liveDrive.name || site.default_library_name }
                                    : site),
                                });
                              }
                            } catch {
                              // Keep the configured drive as a fallback if live resolution is unavailable.
                            }
                          }
                          // folderPathStack[3+] were built from vessel.vessel_folder_path
                          // segments with placeholder ids (a DB vessel id or a
                          // "<segment>-<n>" string) — neither is a real Graph
                          // drive-item id. Folder view's live children fetch is
                          // keyed by real id, so walk the actual site/drive tree
                          // from root, matching each segment by name (vessel-
                          // prefix-insensitive), and swap in the real ids we find.
                          // Without this, "View Documents" lands on a folder id
                          // Graph has never heard of and the listing (and every
                          // breadcrumb slice of it) renders empty.
                          if (navSite?.site_id && resolvedDriveId && folderPathStack[0]?.id === 'sites_root' && folderPathStack.length > 3) {
                            const resolveSiteId = navSite.site_id;
                            const resolveDriveId = resolvedDriveId;
                            try {
                              const stripPrefix = (s: string) => s.replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '');
                              const resolvedPath = folderPathStack.slice(0, 3);
                              let parentId = 'root';
                              for (let i = 3; i < folderPathStack.length; i++) {
                                const segName = folderPathStack[i].name.trim().toLowerCase();
                                // eslint-disable-next-line no-await-in-loop
                                const children = await host._loadAndCacheSiteFolderChildren(resolveSiteId, resolveDriveId, parentId);
                                const match = (children || []).find((item: any) => !!item.folder && (
                                  (item.name || '').trim().toLowerCase() === segName ||
                                  stripPrefix((item.name || '').trim().toLowerCase()) === stripPrefix(segName)
                                ));
                                if (!match?.id) break;
                                resolvedPath.push({ id: match.id, name: match.name || folderPathStack[i].name });
                                parentId = match.id;
                              }
                              if (resolvedPath.length > 3) {
                                folderPathStack = [...resolvedPath, ...folderPathStack.slice(resolvedPath.length)];
                              }
                            } catch {
                              // Keep placeholder ids; Folder view will show an empty state for the unresolved tail.
                            }
                          }
                          host._pushFolderNav(folderPathStack, nav.docMainFolder);
                          host.setState({
                            view: 'list',
                            docViewMode: 'folder',
                            vesselFilter: nav.vesselFilter,
                            docScopeType: folderPathStack[0]?.id === 'sites_root' ? 'sites' : 'vessels',
                            activeDocumentSite: navSite?.site_key || host.state.activeDocumentSite,
                            docMainFolder: nav.docMainFolder,
                            folderPathStack,
                            docGroupFilter: 'all',
                            docCategoryFilter: 'all',
                            docGroupLevelFilter: 'all',
                            docLeafCategoryFilter: 'all',
                            docSubCategoryFilter: 'all',
                            catFilter: 'all',
                            docListPage: 0,
                          });
                          if (folderPathStack[0]?.id !== 'sites_root') {
                            void host._loadFilesForVessel(vessel.name).catch(() => undefined);
                          }
                        }}
                        title={isSharePointOnly ? 'Confirm this vessel above to view its documents' : undefined}
                        style={{
                          flex: 1, background: isSharePointOnly ? clay.surface : clay.accentGradient,
                          color: isSharePointOnly ? clay.textMuted : '#fff',
                          border: 'none', borderRadius: clay.radiusIcon, padding: '8px 10px',
                          fontSize: 12, fontWeight: 700, cursor: isSharePointOnly ? 'not-allowed' : 'pointer',
                          opacity: isSharePointOnly ? 0.7 : 1,
                          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                          boxShadow: isSharePointOnly ? 'none' : clay.shadowButton,
                        }}
                      >
                        📄 View Documents
                      </button>
                      <a
                        href={getSpoVesselFolderUrl(host, vessel)}
                        target="_blank"
                        rel="noopener noreferrer"
                        onClick={e => e.stopPropagation()}
                        title="Open in SharePoint"
                        style={{
                          width: 36, display: 'flex', alignItems: 'center', justifyContent: 'center',
                          border: 'none', borderRadius: clay.radiusIcon, background: clay.surface,
                          color: clay.accentDark, fontSize: 14, textDecoration: 'none',
                          boxShadow: clay.shadowRaised,
                        }}
                      >
                        ↗
                      </a>
                    </div>

                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* ── Normal Folders Section ── */}
        {renderNormalFoldersSection(host)}

        {/* Keyframe animation for spinners */}
        <style>{`
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

          /* Vessel card banner: a restrained, professional zoom on hover —
             no cartoon bob/sway/drift animation, just a subtle Ken-Burns-style
             scale so the card feels alive without looking playful. */
          .vessel-card .vessel-card-image {
            transition: transform 0.4s ease;
            transform: scale(1.0);
            will-change: transform;
          }

          .vessel-card:hover .vessel-card-image {
            transform: scale(1.045);
          }

          @media (prefers-reduced-motion: reduce) {
            .vessel-card .vessel-card-image {
              transition: none !important;
              transform: none !important;
            }
          }
        `}</style>

        {/* ── Dismiss Confirm Dialog ── */}
        {renderDismissConfirmDialog(host)}

        {/* ── Provision Dialog ── */}
        {renderProvisionDialog(host)}

        {/* ── Vessel Field Save Confirm Dialog ── */}
        {renderVesselFieldConfirmDialog(host)}

        {/* Modals */}
        {modal === 'create' && host._renderVesselForm('create')}
        {modal === 'edit' && host._renderVesselForm('edit')}
        {host.state.vesselActionPicker && renderVesselActionPicker(host)}
        {modal === 'delete' && host._renderDeleteModal()}
      </div>
    );
}
