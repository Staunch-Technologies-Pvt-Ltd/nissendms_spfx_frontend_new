import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { OcrStagingItem, DocumentCategory, TagFieldDef } from '../types/ui';
import { isMobileWidth, isTabletWidth } from '../responsive';
import { clay } from '../clayTheme';

// ── Production Term Store Taxonomy ───────────────────────────────────────────
export const PRODUCTION_VESSELS: string[] = [
  "Belle Lune", "Bow Fighter", "Bow Fraternity", "Cameroun Express", "Cecilie F",
  "Cote D Ivoire Express", "Dutches Emerald", "Ghana Express", "Lignum Grid",
  "Lignum Mesh", "Lignum Web", "Maersk EI Banco", "Maersk EI Palomar",
  "Maersk Ferrato", "Maersk Finisterre", "Maersk Frio", "Norse Evolution",
  "Norse Ijmuiden", "Norse New Haven", "Peissy", "Potiniere", "Senegal Express",
  "Snow Flake", "Snow Flower"
];

export const DRAWING_TAXONOMY: Record<string, string[]> = {
  "Basic": [
    "Basic Drawings", "Capacity Plan & Dead Weight", "Damage Control Plan",
    "Docking Plan", "EEDI Technical file", "Emergency Towing Booklet",
    "General Arrangement", "Loading Manual", "Ship Structure Access Manuals",
    "Trim & Stability Information"
  ],
  "Electrical": [
    "Electrical Drawings", "Emergency Switchboard Arrangement",
    "Main Switchboard Arrangement", "Power Distribution Diagram", "Single Line Diagram"
  ],
  "Hull": [
    "Bulkhead plans", "Cargo Securing Manual", "Container Stowage Plan",
    "Hull Drawings", "Makers List of Hull parts", "Midship Section",
    "Mooring Arrangement", "Painting Schedule", "Profile & Deck Plan",
    "Results of Official Sea Trial", "Rudder and Rudder stock",
    "Shell Expansion", "Superstructure"
  ],
  "Machinery": [
    "Arrangement of Engine Room", "Machinery Drawings", "Machinery Makers List",
    "Machinery Particulars", "Pipeline Diagram", "Shafting Arrangements",
    "Stern Tube", "Test Record of Official Sea Trial", "Other Drawings"
  ],
  "Safety": [
    "Fire Control Plan", "Life Saving Appliances Plan", "Safety Drawings"
  ],
  "Archive": [
    "Archive"
  ]
};

export const MANUAL_TAXONOMY: Record<string, string[]> = {
  "Automation": ["Alarm Monitoring System", "Engine Control System"],
  "Auxiliary Engine": ["Operation & Maintenance Manual"],
  "Boiler": ["Operation & Maintenance Manual"],
  "Bridge Equipments": ["Bridge Equipments"],
  "Cargo": [
    "Ballast System Manual", "Cargo Crane Manual", "Cargo Pump Manual",
    "COW Manual", "Hatch Cover Manual", "IG System Manual", "ODME Manual"
  ],
  "Deck Machinery": ["Mooring Winch Manual", "Windlass Manual"],
  "Electrical": ["Main Switchboard Manual", "Power Management System"],
  "Main Engine": ["Operation & Maintenance Manual", "Other Manuals"],
  "Pollution": [
    "BWTS Manual", "EGR System Manual", "Exhaust Gas Scrubber Manual",
    "Incinerator Manual", "OWS Manual", "SCR System Manual", "Sewage Treatment Plant Manual"
  ],
  "Propulsion": ["Shaft Generator Manual"],
  "Refrigeration": ["AC Plant Manual"],
  "Safety": [
    "CO2 System Manual", "Emergency Generator Manual",
    "Fire Alarm Manual", "Fire Detection System Manual"
  ],
  "Shafting": ["CPP Manual", "Stern Tube Manual"],
  "Steering Gear": ["Maintenance Manual", "Operation Manual"],
  "Thrusters": ["Operation & Maintenance Manual"],
  "To Be Classified": ["To Be Classified"]
};

// ── Confidence Floor ─────────────────────────────────────────────────────────
export const VESSEL_CONFIDENCE_FLOOR = 0.60;

// Helper to safely extract field values and confidence
export function getFieldData(tags: any, fieldKey: string): { value: string; confidence: number; tier: number } {
  if (!tags) return { value: '', confidence: 0, tier: 2 };
  const raw = tags[fieldKey];
  if (raw && typeof raw === 'object' && 'value' in raw) {
    const conf = typeof raw.confidence === 'number' ? raw.confidence : 0;
    const rawVal = String(raw.value || '');
    // If field is 'vessel' and confidence is below floor (0.60), treat value as blank
    const val = (fieldKey === 'vessel' && conf < VESSEL_CONFIDENCE_FLOOR) ? '' : rawVal;
    return {
      value: val,
      confidence: conf,
      tier: raw.tier === 1 ? 1 : 2,
    };
  }
  if (typeof raw === 'string') {
    return { value: raw, confidence: 0.88, tier: 1 };
  }
  return { value: '', confidence: 0, tier: 2 };
}

// Format timestamp for OCR scan time display
export function formatScannedTime(ts?: string | number | null): string {
  if (!ts) return 'Just now';
  try {
    const d = typeof ts === 'number' ? new Date(ts) : new Date(String(ts));
    if (isNaN(d.getTime())) return String(ts);
    return d.toLocaleString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  } catch {
    return String(ts);
  }
}

// Small badge for rendering per-field confidence and tier
export const FieldConfidenceBadge: React.FC<{ confidence: number; tier?: number }> = ({ confidence, tier }) => {
  if (confidence <= 0) return null;
  const pct = Math.round(confidence * 100);
  const isHigh = confidence >= 0.85;
  const isMed = confidence >= 0.60 && confidence < 0.85;

  const bg = isHigh ? '#dcfce7' : isMed ? '#fef3c7' : '#fee2e2';
  const color = isHigh ? '#15803d' : isMed ? '#b45309' : '#dc2626';
  const border = isHigh ? '#86efac' : isMed ? '#fde68a' : '#fca5a5';

  return (
    <span
      title={`Tier ${tier || 2} Extraction: ${pct}% Confidence (${isHigh ? 'High Confidence' : 'Review Recommended'})`}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 3,
        fontSize: 10,
        fontWeight: 700,
        padding: '1px 5px',
        borderRadius: 4,
        background: bg,
        color: color,
        border: `1px solid ${border}`,
        marginLeft: 4,
        verticalAlign: 'middle',
      }}
    >
      <span style={{ fontSize: 6 }}>●</span>
      {pct}%
    </span>
  );
};

interface OcrStagingQueueProps {
  host: VesselEmail;
  categories: DocumentCategory[];
  onRefreshCategories: () => void;
}

export const OcrStagingQueue: React.FC<OcrStagingQueueProps> = ({ host, categories, onRefreshCategories }) => {
  const [items, setItems] = React.useState<OcrStagingItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);

  // Queue tabs: 'staged' = ready-for-review, 'unstaged' = needs_review (unclassifiable)
  const [activeQueueTab, setActiveQueueTab] = React.useState<'staged' | 'unstaged'>('staged');

  // Filters
  const [filterSource, setFilterSource] = React.useState<'all' | 'folder' | 'direct'>('all');
  const [searchQuery, setSearchQuery] = React.useState<string>('');

  // Edit dialog state
  const [editingItem, setEditingItem] = React.useState<OcrStagingItem | null>(null);
  const [selectedCatId, setSelectedCatId] = React.useState<number | null>(null);
  const [editTags, setEditTags] = React.useState<Record<string, any>>({});
  const [isSavingEdit, setIsSavingEdit] = React.useState<boolean>(false);
  const [editError, setEditError] = React.useState<string | null>(null);

  // ── "Which vessel is this?" Single-Item Action Prompt State (Option A) ──────
  const [vesselPromptItem, setVesselPromptItem] = React.useState<{ item: OcrStagingItem; action: 'promote' | 'move' } | null>(null);
  const [promptVesselChoice, setPromptVesselChoice] = React.useState<string>('');
  const [isCreatingNewVesselInline, setIsCreatingNewVesselInline] = React.useState<boolean>(false);
  const [newVesselInput, setNewVesselInput] = React.useState<string>('');
  const [isSavingPrompt, setIsSavingPrompt] = React.useState<boolean>(false);
  const [promptError, setPromptError] = React.useState<string | null>(null);

  // ── Batch Vessel Assignment State for Needs Review Tab (Option 4b) ──────────
  const [selectedUnstagedIds, setSelectedUnstagedIds] = React.useState<number[]>([]);
  const [batchModalOpen, setBatchModalOpen] = React.useState<boolean>(false);
  const [batchVesselChoice, setBatchVesselChoice] = React.useState<string>('');
  const [batchPromoteImmediately, setBatchPromoteImmediately] = React.useState<boolean>(true);
  const [isApplyingBatch, setIsApplyingBatch] = React.useState<boolean>(false);
  const [batchError, setBatchError] = React.useState<string | null>(null);

  // Promoting unstaged -> staged
  const [promotingIds, setPromotingIds] = React.useState<Record<number, boolean>>({});

  // Action loading state (item id -> boolean)
  const [movingIds, setMovingIds] = React.useState<Record<number, boolean>>({});
  const [actionSuccess, setActionSuccess] = React.useState<{ id: number; message: string; path?: string; warning?: string } | null>(null);

  const fetchQueue = React.useCallback(async (showSpinner: boolean = false) => {
    if (showSpinner) setLoading(true);
    setError(null);
    try {
      const base = host._base();
      const res = await fetch(`${base}/api/ocr/staging`, { headers: host._headers() });
      if (!res.ok) {
        throw new Error(`Failed to load staging queue (status ${res.status})`);
      }
      const data: OcrStagingItem[] = await res.json();
      setItems(data);
      void host._syncPendingVesselSuggestionsFromStaging(data, host.state.view === 'templates');
      // Count only staged items for the badge
      const stagedCount = data.filter(i => i.status !== 'needs_review').length;
      if (host.state.ocrStagingCount !== stagedCount) {
        host.setState({ ocrStagingCount: stagedCount });
      }
    } catch (err: any) {
      setError(err?.message || 'Error fetching staging items');
    } finally {
      if (showSpinner) setLoading(false);
    }
  }, [host]);

  // Initial load & polling every 8s
  React.useEffect(() => {
    void fetchQueue(true);
    const interval = window.setInterval(() => {
      void fetchQueue(false);
    }, 8000);
    return () => window.clearInterval(interval);
  }, [fetchQueue]);

  React.useEffect(() => {
    const hint = host.state.ocrQueueTabHint;
    if (!hint) return;
    setActiveQueueTab(hint);
    host.setState({ ocrQueueTabHint: null });
  }, [host, host.state.ocrQueueTabHint]);

  // Open Edit Dialog
  const handleOpenEdit = (item: OcrStagingItem) => {
    setEditingItem(item);
    const initialCatId = item.category_id || (categories.length > 0 ? categories[0].id : null);
    setSelectedCatId(initialCatId);
    
    // Normalize edit tags from either per-field objects or flat strings
    const rawTags = item.suggested_tags || {};
    const initialTags: Record<string, string> = {};
    for (const key of ['category', 'group', 'sub_category', 'vessel', 'department']) {
      const fieldData = getFieldData(rawTags, key);
      initialTags[key] = fieldData.value;
    }
    const vData = getFieldData(rawTags, 'vessel');
    if (!initialTags.category && initialCatId) {
      const cat = categories.find(c => c.id === initialCatId);
      if (cat) initialTags.category = cat.name;
    }
    setEditTags(initialTags);
    setEditError(null);
  };

  // Active category definition in edit modal
  const activeCategory = React.useMemo(() => {
    return categories.find(c => c.id === selectedCatId) || null;
  }, [categories, selectedCatId]);

  // Resolved dynamic path preview in edit modal: folder ends at Category level
  const livePreviewPath = React.useMemo(() => {
    if (!activeCategory) return '';
    let templateStr = activeCategory.dms_path_template || 'Technical & Crewing/{vessel}/Drawings and Manuals/{group}/{category}';
    templateStr = templateStr.replace(/\/\{sub_?category\}/gi, '');
    let path = templateStr;
    for (const [k, v] of Object.entries(editTags)) {
      path = path.replace(new RegExp(`\\{${k}\\}`, 'g'), String(v || '').trim());
    }
    path = path.replace(/\{[a-z0-9_]+\}/g, '');
    path = path.replace(/\/+/g, '/').replace(/^\/+|\/+$/g, '');
    return path;
  }, [activeCategory, editTags]);

  // Save edited tags
  const handleSaveTags = async () => {
    if (!editingItem) return;
    setIsSavingEdit(true);
    setEditError(null);
    try {
      const base = host._base();
      const res = await fetch(`${base}/api/ocr/staging/${editingItem.id}/tags`, {
        method: 'PATCH',
        headers: host._headers(),
        body: JSON.stringify({
          category_id: selectedCatId,
          suggested_tags: editTags,
        }),
      });
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.detail || `Failed to update tags (status ${res.status})`);
      }
      setEditingItem(null);
      void fetchQueue(false);
    } catch (err: any) {
      setEditError(err?.message || 'Failed to save tags.');
    } finally {
      setIsSavingEdit(false);
    }
  };

  // Move to SharePoint (with prompt if vessel is missing)
  const handleMoveItem = async (item: OcrStagingItem, overrideVessel?: string) => {
    const vData = getFieldData(item.suggested_tags, 'vessel');
    const vesselVal = overrideVessel || (vData.confidence >= VESSEL_CONFIDENCE_FLOOR ? vData.value : '');

    // If vessel is undetected/blank, prompt user first (Option A)
    if (!vesselVal) {
      setPromptVesselChoice('');
      setPromptError(null);
      setIsCreatingNewVesselInline(false);
      setNewVesselInput('');
      setVesselPromptItem({ item, action: 'move' });
      return;
    }

    setMovingIds(prev => ({ ...prev, [item.id]: true }));
    setActionSuccess(null);
    try {
      const base = host._base();
      const res = await fetch(`${base}/api/ocr/staging/${item.id}/move`, {
        method: 'POST',
        headers: host._headers(),
      });
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.detail || `Move failed with status ${res.status}`);
      }
      const data = await res.json();
      const metadataPatch = data?.metadata_patch;
      const metadataWarning = metadataPatch
        ? metadataPatch.partial_success
          ? `Metadata patch warning: ${metadataPatch.vessel_rest_error?.error || metadataPatch.warning || metadataPatch.error || metadataPatch.reason || 'Some OCR tags were not saved.'}`
          : metadataPatch.ok === false
            ? `Metadata patch warning: ${metadataPatch.error || metadataPatch.reason || 'Columns were not updated.'}`
            : undefined
        : undefined;

      setActionSuccess({
        id: item.id,
        message: `Successfully filed "${item.filename}" into SharePoint Online!`,
        path: data.target_path,
        warning: metadataWarning,
      });
      // Refresh vessel files in state
      if (vesselVal) {
        void host._mergeLiveSharePointFiles([vesselVal]).catch(() => undefined);
      }
      void fetchQueue(false);
    } catch (err: any) {
      alert(`Move error: ${err?.message || 'Could not relocate file in SharePoint'}`);
    } finally {
      setMovingIds(prev => ({ ...prev, [item.id]: false }));
    }
  };

  // Promote a needs_review item to tag_suggested (with prompt if vessel is missing)
  const handlePromoteToStaged = async (item: OcrStagingItem, overrideVessel?: string) => {
    const vData = getFieldData(item.suggested_tags, 'vessel');
    const vesselVal = overrideVessel || (vData.confidence >= VESSEL_CONFIDENCE_FLOOR ? vData.value : '') || '';

    // If vessel is undetected/blank, prompt user first (Option A)
    if (!vesselVal) {
      setPromptVesselChoice('');
      setPromptError(null);
      setIsCreatingNewVesselInline(false);
      setNewVesselInput('');
      setVesselPromptItem({ item, action: 'promote' });
      return;
    }

    setPromotingIds(prev => ({ ...prev, [item.id]: true }));
    try {
      const base = host._base();
      const res = await fetch(`${base}/api/ocr/staging/${item.id}/stage`, {
        method: 'POST',
        headers: host._headers(),
      });
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.detail || `Promote failed (status ${res.status})`);
      }
      setActiveQueueTab('staged');
      void fetchQueue(false);
    } catch (err: any) {
      alert(`Could not promote to staging: ${err?.message}`);
    } finally {
      setPromotingIds(prev => ({ ...prev, [item.id]: false }));
    }
  };

  // Confirm single-item vessel prompt (Option A)
  const handleConfirmVesselPrompt = async () => {
    if (!vesselPromptItem) return;
    let vesselToSet = promptVesselChoice.trim();
    if (isCreatingNewVesselInline) {
      vesselToSet = newVesselInput.trim();
      if (!vesselToSet) {
        setPromptError('Please enter a vessel name.');
        return;
      }
      try {
        const base = host._base();
        const createRes = await fetch(`${base}/api/vessels`, {
          method: 'POST',
          headers: host._headers(),
          body: JSON.stringify({ name: vesselToSet, vessel_type: 'Other' }),
        });
        if (createRes.ok) {
          const createdVessel = await createRes.json();
          if (createdVessel && createdVessel.id) {
            host.setState(prev => ({
              vessels: [...(prev.vessels || []), createdVessel],
            }));
          }
        }
      } catch (err) {
        // Proceed even if cache sync delay
      }
    }
    if (!vesselToSet) {
      setPromptError('Please select or enter a vessel name.');
      return;
    }

    setIsSavingPrompt(true);
    setPromptError(null);
    try {
      const base = host._base();
      const rawTags = vesselPromptItem.item.suggested_tags || {};
      const updatedTags: Record<string, any> = {};
      for (const k of ['group', 'category', 'sub_category', 'department']) {
        updatedTags[k] = getFieldData(rawTags, k).value;
      }
      updatedTags.vessel = vesselToSet;

      const patchRes = await fetch(`${base}/api/ocr/staging/${vesselPromptItem.item.id}/tags`, {
        method: 'PATCH',
        headers: host._headers(),
        body: JSON.stringify({
          category_id: vesselPromptItem.item.category_id,
          suggested_tags: updatedTags,
        }),
      });
      if (!patchRes.ok) {
        const err = await patchRes.json().catch(() => ({}));
        throw new Error(err.detail || `Failed to assign vessel (status ${patchRes.status})`);
      }

      const currentItem = vesselPromptItem.item;
      const action = vesselPromptItem.action;
      setVesselPromptItem(null);

      if (action === 'promote') {
        await handlePromoteToStaged(currentItem, vesselToSet);
      } else if (action === 'move') {
        await handleMoveItem(currentItem, vesselToSet);
      } else {
        void fetchQueue(false);
      }
    } catch (err: any) {
      setPromptError(err?.message || 'Error assigning vessel.');
    } finally {
      setIsSavingPrompt(false);
    }
  };

  // Apply batch vessel assignment to selected Needs Review items (Option 4b)
  const handleApplyBatchVessel = async () => {
    if (!selectedUnstagedIds.length) return;
    const vesselToSet = batchVesselChoice.trim();
    if (!vesselToSet) {
      setBatchError('Please select a vessel name.');
      return;
    }
    setIsApplyingBatch(true);
    setBatchError(null);
    try {
      const base = host._base();
      for (const itemId of selectedUnstagedIds) {
        const item = items.find(i => i.id === itemId);
        if (!item) continue;
        const rawTags = item.suggested_tags || {};
        const updatedTags: Record<string, any> = {};
        for (const k of ['group', 'category', 'sub_category', 'department']) {
          updatedTags[k] = getFieldData(rawTags, k).value;
        }
        updatedTags.vessel = vesselToSet;

        await fetch(`${base}/api/ocr/staging/${itemId}/tags`, {
          method: 'PATCH',
          headers: host._headers(),
          body: JSON.stringify({
            category_id: item.category_id,
            suggested_tags: updatedTags,
          }),
        });

        if (batchPromoteImmediately) {
          await fetch(`${base}/api/ocr/staging/${itemId}/stage`, {
            method: 'POST',
            headers: host._headers(),
          });
        }
      }

      setBatchModalOpen(false);
      setSelectedUnstagedIds([]);
      if (batchPromoteImmediately) {
        setActiveQueueTab('staged');
      }
      void fetchQueue(false);
    } catch (err: any) {
      setBatchError(err?.message || 'Failed to apply batch vessel assignment.');
    } finally {
      setIsApplyingBatch(false);
    }
  };

  // Dismiss item
  const handleDismissItem = async (item: OcrStagingItem) => {
    if (!window.confirm(`Dismiss "${item.filename}" from the review queue?`)) return;
    try {
      const base = host._base();
      await fetch(`${base}/api/ocr/staging/${item.id}`, {
        method: 'DELETE',
        headers: host._headers(),
      });
      void fetchQueue(false);
    } catch (err: any) {
      alert(`Failed to dismiss: ${err?.message}`);
    }
  };

  // Split all items into staged and unstaged
  const stagedItems = React.useMemo(() => {
    return items.filter(item => item.status !== 'needs_review');
  }, [items]);

  const unstagedItems = React.useMemo(() => {
    return items.filter(item => item.status === 'needs_review');
  }, [items]);

  // Apply source + search filters
  const applyFilters = (list: OcrStagingItem[]) => list.filter(item => {
    if (filterSource !== 'all' && item.upload_source !== filterSource) return false;
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      const matchName = item.filename.toLowerCase().includes(q);
      const matchVessel = (item.vessel_name || '').toLowerCase().includes(q);
      const matchCat = (item.category_name || '').toLowerCase().includes(q);
      if (!matchName && !matchVessel && !matchCat) return false;
    }
    return true;
  });

  const filteredStaged = React.useMemo(() => applyFilters(stagedItems), [stagedItems, filterSource, searchQuery]);
  const filteredUnstaged = React.useMemo(() => applyFilters(unstagedItems), [unstagedItems, filterSource, searchQuery]);
  // Backward compat — filteredItems is whichever tab is active
  const filteredItems = activeQueueTab === 'staged' ? filteredStaged : filteredUnstaged;
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const isMobile = isMobileWidth(viewportWidth);
  const isTablet = isTabletWidth(viewportWidth);
  const unidentifiedNamesLower = React.useMemo(
    () => new Set((host.state.ocrUnidentifiedFiles || []).map(n => (n || '').trim().toLowerCase()).filter(Boolean)),
    [host.state.ocrUnidentifiedFiles]
  );
  const unidentifiedRows = React.useMemo(
    () => filteredUnstaged.filter(item => unidentifiedNamesLower.has((item.filename || '').trim().toLowerCase())),
    [filteredUnstaged, unidentifiedNamesLower]
  );
  const unidentifiedFileNames = React.useMemo(
    () => Array.from(new Set(unidentifiedRows.map(item => (item.filename || '').trim()).filter(Boolean))),
    [unidentifiedRows]
  );

  const vesselOptions = React.useMemo(() => {
    const list = [...PRODUCTION_VESSELS, ...(host.state.vessels || []).map(v => v.name).filter(Boolean)];
    return Array.from(new Set(list)).sort();
  }, [host.state.vessels]);

  // Derived list of unstaged files that have no vessel assigned
  const unstagedMissingVessel = React.useMemo(() => {
    return filteredUnstaged.filter(item => {
      const vData = getFieldData(item.suggested_tags, 'vessel');
      return !vData.value && !item.vessel_name;
    });
  }, [filteredUnstaged]);

  const pendingCreateSuggestions = React.useMemo(
    () => (host.state.pendingVesselSuggestions || []).filter(s => !s.matchedExisting),
    [host.state.pendingVesselSuggestions]
  );

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      {pendingCreateSuggestions.length > 0 && (
        <div style={{
          background: 'linear-gradient(135deg, #ecfdf5 0%, #ebf9ff 100%)',
          border: `1px solid ${clay.accentSoft}`,
          borderRadius: clay.radiusCard,
          padding: '14px 18px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 12,
          flexWrap: 'wrap',
          boxShadow: clay.shadowRaised,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
            <span style={{ fontSize: 18 }}>✨</span>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: '#14532d' }}>
                New vessel suggestion{pendingCreateSuggestions.length > 1 ? 's' : ''} detected from OCR
              </div>
              <div style={{ fontSize: 12, color: '#166534', marginTop: 2, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {pendingCreateSuggestions.length} candidate{pendingCreateSuggestions.length > 1 ? 's' : ''}: {pendingCreateSuggestions.slice(0, 3).map(s => s.vesselName || 'Unnamed').join(', ')}{pendingCreateSuggestions.length > 3 ? ' ...' : ''}
              </div>
            </div>
          </div>
          <button
            type="button"
            onClick={() => host._openSuggestedVesselsFromSidebar()}
            style={{
              padding: '8px 14px',
              borderRadius: 10,
              border: 'none',
              background: clay.accentGradient,
              color: '#fff',
              fontSize: 12,
              fontWeight: 800,
              cursor: 'pointer',
              whiteSpace: 'nowrap',
              boxShadow: clay.shadowButton,
              transition: 'all 0.15s ease',
            }}
            onMouseEnter={e => { e.currentTarget.style.background = clay.accentGradientHover; e.currentTarget.style.boxShadow = clay.shadowButton.replace('rgba(221,145,89,0.35)', 'rgba(199,122,62,0.42)'); }}
            onMouseLeave={e => { e.currentTarget.style.background = clay.accentGradient; e.currentTarget.style.boxShadow = clay.shadowButton; }}
          >
            Review / Create Vessel
          </button>
        </div>
      )}

      {/* Staged / Unstaged Tab Strip */}
      <div style={{
        display: 'flex',
        gap: 8,
        padding: '8px 10px 0',
        borderBottom: `2px solid ${clay.accentSoft}`,
      }}>
        {[
          { key: 'staged' as const, label: '📋 Staging Queue', count: stagedItems.length, color: clay.accentDark, bg: clay.accentSoft },
          { key: 'unstaged' as const, label: '⚠️ Needs Review', count: unstagedItems.length, color: '#b45309', bg: '#fef3c7' },
        ].map(tab => {
          const isActive = activeQueueTab === tab.key;
          return (
            <button
              key={tab.key}
              type="button"
              onClick={() => {
                setActiveQueueTab(tab.key);
                setSelectedUnstagedIds([]);
              }}
              style={{
                padding: '12px 18px', border: 'none', background: isActive ? clay.surfaceRaised : 'transparent', cursor: 'pointer',
                fontSize: 13, fontWeight: 800,
                color: isActive ? tab.color : '#64748b',
                borderBottom: isActive ? `3px solid ${tab.color}` : '3px solid transparent',
                marginBottom: -2, transition: 'all 0.15s',
                display: 'flex', alignItems: 'center', gap: 8,
                borderRadius: '14px 14px 0 0',
                boxShadow: isActive ? 'inset 0 1px 0 rgba(255,255,255,0.8)' : 'none',
              }}
            >
              {tab.label}
              <span style={{
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                minWidth: 22, height: 22, borderRadius: 11,
                background: tab.count > 0 ? (isActive ? tab.color : tab.bg) : '#f1f5f9',
                color: tab.count > 0 ? (isActive ? '#fff' : tab.color) : '#94a3b8',
                fontSize: 11, fontWeight: 800,
                boxShadow: 'inset 0 1px 2px rgba(255,255,255,0.7)',
              }}>
                {tab.count}
              </span>
            </button>
          );
        })}
      </div>

      {/* Unstaged area notice */}
      {activeQueueTab === 'unstaged' && (
        <div style={{
          background: '#fef9c3', border: '1px solid #fde047', borderRadius: 10, padding: '12px 16px',
          fontSize: 13, color: '#713f12', display: 'flex', alignItems: 'flex-start', gap: 10,
        }}>
          <span style={{ fontSize: 18, flexShrink: 0 }}>⚠️</span>
          <div>
            <strong>Unstaged files — manual classification required.</strong>
            <div style={{ marginTop: 4, lineHeight: 1.5 }}>
              These files could not be automatically classified because the vessel could not be identified, text extraction failed, or overall confidence was below the minimum threshold.
              Use the <strong>Edit</strong> button to set the correct Vessel, Group, Category and Sub-Category, or assign vessels in bulk below, then click <strong>Move to Staging</strong> to promote the file to the review queue.
            </div>
          </div>
        </div>
      )}

      {activeQueueTab === 'unstaged' && unidentifiedFileNames.length > 0 && (
        <div style={{
          background: '#fff7ed', border: '1px solid #fdba74', borderRadius: 10, padding: '12px 16px',
          fontSize: 13, color: '#9a3412',
        }}>
          <div style={{ fontWeight: 800, marginBottom: 8 }}>Vessel name not identified files</div>
          <div style={{ fontSize: 12, marginBottom: 8 }}>
            These uploaded files could not be mapped to a vessel during OCR extraction and are listed here for manual review.
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {unidentifiedFileNames.slice(0, 8).map(name => (
              <div key={name} style={{ fontSize: 12, color: '#7c2d12', wordBreak: 'break-word' }}>
                • {name}
              </div>
            ))}
            {unidentifiedFileNames.length > 8 && (
              <div style={{ fontSize: 12, color: '#9a3412' }}>+{unidentifiedFileNames.length - 8} more file(s)</div>
            )}
          </div>
        </div>
      )}

      {/* ── Sticky Batch Vessel Assignment Toolbar on Needs Review Tab (Option 4b) ── */}
      {activeQueueTab === 'unstaged' && filteredUnstaged.length > 0 && (
        <div style={{
          background: 'linear-gradient(90deg, #f0fdf4 0%, #eff6ff 100%)',
          border: '1px solid #bfdbfe', borderRadius: 12, padding: '12px 18px',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12,
          boxShadow: '0 2px 6px rgba(0,0,0,0.03)',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <span style={{ fontSize: 18 }}>🚢</span>
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: '#1e3a8a' }}>
                {selectedUnstagedIds.length > 0 ? (
                  <span>✓ {selectedUnstagedIds.length} of {filteredUnstaged.length} file{filteredUnstaged.length > 1 ? 's' : ''} selected</span>
                ) : unstagedMissingVessel.length > 0 ? (
                  <span>{unstagedMissingVessel.length} file{unstagedMissingVessel.length > 1 ? 's' : ''} missing vessel assignment</span>
                ) : (
                  <span>Bulk Vessel Assignment</span>
                )}
              </div>
              <div style={{ fontSize: 11, color: '#64748b', marginTop: 1 }}>
                Select files to assign a vessel in bulk and optionally promote directly to the Staging Queue
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            {unstagedMissingVessel.length > 0 && (
              <button
                type="button"
                onClick={() => {
                  setSelectedUnstagedIds(unstagedMissingVessel.map(i => i.id));
                }}
                style={{
                  padding: '7px 12px', background: '#fff', border: '1px solid #93c5fd', borderRadius: 8,
                  fontSize: 12, fontWeight: 700, color: '#0369a1', cursor: 'pointer',
                  transition: 'all 0.15s',
                }}
              >
                Select All Missing Vessel ({unstagedMissingVessel.length})
              </button>
            )}

            {selectedUnstagedIds.length > 0 ? (
              <>
                <button
                  type="button"
                  onClick={() => {
                    setBatchVesselChoice('');
                    setBatchError(null);
                    setBatchModalOpen(true);
                  }}
                  style={{
                    padding: '7px 16px', background: clay.accentGradient, border: 'none', borderRadius: 8,
                    fontSize: 12, fontWeight: 800, color: '#fff', cursor: 'pointer',
                    boxShadow: clay.shadowButton,
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                    transition: 'all 0.15s ease',
                  }}
                  onMouseEnter={e => { e.currentTarget.style.background = clay.accentGradientHover; e.currentTarget.style.boxShadow = clay.shadowButton.replace('rgba(221,145,89,0.35)', 'rgba(199,122,62,0.42)'); }}
                  onMouseLeave={e => { e.currentTarget.style.background = clay.accentGradient; e.currentTarget.style.boxShadow = clay.shadowButton; }}
                >
                  <span>🚢 Assign Vessel ({selectedUnstagedIds.length})</span>
                </button>
                <button
                  type="button"
                  onClick={() => setSelectedUnstagedIds([])}
                  style={{
                    padding: '7px 10px', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8,
                    fontSize: 12, fontWeight: 600, color: '#64748b', cursor: 'pointer',
                  }}
                >
                  Clear Selection
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={() => setSelectedUnstagedIds(filteredUnstaged.map(i => i.id))}
                style={{
                  padding: '7px 12px', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8,
                  fontSize: 12, fontWeight: 600, color: '#475569', cursor: 'pointer',
                }}
              >
                Select All ({filteredUnstaged.length})
              </button>
            )}
          </div>
        </div>
      )}

      {/* Top Banner / Filter Bar */}
      <div style={{
        background: 'linear-gradient(135deg, rgba(240,251,250,0.96), rgba(232,246,245,0.96))', borderRadius: clay.radiusCard, padding: '16px 20px', border: `1px solid ${clay.accentSoft}`,
        display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 14,
        boxShadow: clay.shadowRaised,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: clay.textMuted }}>Filter Source:</span>
          {(['all', 'folder', 'direct'] as const).map(src => (
            <button
              key={src}
              type="button"
              onClick={() => setFilterSource(src)}
              style={{
                padding: '6px 14px', borderRadius: 20, border: '1px solid',
                borderColor: filterSource === src ? clay.accent : clay.accentSoft,
                background: filterSource === src ? clay.accentSoft : clay.surface,
                color: filterSource === src ? clay.accentDark : clay.textMuted,
                fontSize: 12, fontWeight: 700, cursor: 'pointer', transition: 'all 0.15s',
              }}
            >
              {src === 'all' ? `All (${filteredItems.length})` : src === 'folder' ? '📁 Folder Uploads' : '📄 Direct Uploads'}
            </button>
          ))}
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 10, width: isMobile ? '100%' : 'auto' }}>
          <input
            type="text"
            placeholder="Search filename or vessel..."
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            style={{
              padding: '7px 12px', borderRadius: 8, border: `1px solid ${clay.accentSoft}`,
              fontSize: 12, width: isMobile ? '100%' : 220, outline: 'none',
            }}
          />
          <button
            type="button"
            onClick={() => void fetchQueue(true)}
            style={{
              padding: '7px 12px', borderRadius: 8, border: `1px solid ${clay.accentSoft}`,
              background: clay.surfaceRaised, color: clay.textMuted, fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            🔄 Refresh
          </button>
        </div>
      </div>

      {/* Success Notification */}
      {actionSuccess && (
        <div style={{
          background: 'linear-gradient(135deg, #ecfdf5 0%, #e0f2fe 100%)', border: `1px solid ${clay.pillActiveBg}`, borderRadius: clay.radiusCard, padding: '12px 16px',
          display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#15803d', fontSize: 13,
          boxShadow: clay.shadowRaised,
        }}>
          <div>
            <strong>✓ {actionSuccess.message}</strong>
            {actionSuccess.path && (
              <div style={{ fontSize: 11, fontFamily: 'monospace', marginTop: 2, color: '#0f766e' }}>
                📁 Destination: {actionSuccess.path}
              </div>
            )}
            {actionSuccess.warning && (
              <div style={{ fontSize: 11, marginTop: 4, color: '#b45309' }}>
                ⚠️ {actionSuccess.warning}
              </div>
            )}
          </div>
          <button
            type="button"
            onClick={() => setActionSuccess(null)}
            style={{ background: 'none', border: 'none', color: '#15803d', cursor: 'pointer', fontWeight: 800 }}
          >
            ✕
          </button>
        </div>
      )}

      {/* Queue Table */}
      <div style={{ background: 'linear-gradient(180deg, #f4fbfb 0%, #eefaf9 100%)', borderRadius: clay.radiusCard, border: `1px solid ${clay.accentSoft}`, overflow: 'hidden', boxShadow: clay.shadowRaised }}>
        {loading && items.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: clay.textMuted, fontSize: 14 }}>
            ⏳ Loading OCR staging queue...
          </div>
        ) : filteredItems.length === 0 ? (
          <div style={{ padding: 48, textAlign: 'center' }}>
            <div style={{ fontSize: 36, marginBottom: 8 }}>🎉</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: clay.text }}>No Documents Pending Review</div>
            <div style={{ fontSize: 13, color: clay.textMuted, marginTop: 4 }}>
              When folders or single files are uploaded, they sit here for AI auto-tagging and user verification.
            </div>
          </div>
        ) : isMobile ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: 10 }}>
            {filteredItems.map(item => {
              const isMoving = !!movingIds[item.id];
              const isPending = item.status === 'ocr_pending';
              const confidence = item.confidence || 0;
              const tags = item.suggested_tags || {};
              const catData = getFieldData(tags, 'category');
              const groupData = getFieldData(tags, 'group');
              const subCatData = getFieldData(tags, 'sub_category');
              const vesselData = getFieldData(tags, 'vessel');
              const catVal = catData.value || item.category_name || 'Drawing';
              const groupVal = groupData.value || '—';
              const subCatVal = subCatData.value || '—';
              // Strictly apply confidence floor to vessel value
              const vesselConfident = vesselData.confidence >= VESSEL_CONFIDENCE_FLOOR;
              const vesselVal = (vesselConfident ? (vesselData.value || item.vessel_name) : '') || '';

              const isSelected = selectedUnstagedIds.includes(item.id);

              return (
                <div key={item.id} style={{ border: isSelected ? `2px solid ${clay.accent}` : `1px solid ${clay.accentSoft}`, borderRadius: 12, padding: 12, background: isSelected ? clay.surfaceRaised : clay.surface }}>
                  <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8 }}>
                    <div style={{ fontSize: 14, fontWeight: 700, color: clay.text, wordBreak: 'break-word' }}>{item.filename}</div>
                    {activeQueueTab === 'unstaged' && (
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={e => {
                          if (e.target.checked) {
                            setSelectedUnstagedIds(prev => [...prev, item.id]);
                          } else {
                            setSelectedUnstagedIds(prev => prev.filter(id => id !== item.id));
                          }
                        }}
                        style={{ cursor: 'pointer', width: 18, height: 18, flexShrink: 0 }}
                      />
                    )}
                  </div>
                  <div style={{ marginTop: 8, display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                    <span style={{ padding: '3px 8px', borderRadius: 12, fontSize: 11, fontWeight: 700, background: confidence >= 0.85 ? '#dcfce7' : confidence >= 0.60 ? '#fef3c7' : '#fee2e2', color: confidence >= 0.85 ? '#15803d' : confidence >= 0.60 ? '#b45309' : '#dc2626' }}>
                      {(confidence * 100).toFixed(0)}% Match
                    </span>
                    <span style={{ padding: '3px 8px', borderRadius: 12, fontSize: 11, fontWeight: 700, background: item.status === 'needs_review' ? '#fee2e2' : '#e0f2fe', color: item.status === 'needs_review' ? '#b91c1c' : '#0369a1' }}>
                      {item.status === 'needs_review' ? 'Needs Review' : item.status === 'ocr_pending' ? 'OCR Pending' : 'Staged'}
                    </span>
                  </div>
                  <div style={{ marginTop: 8, fontSize: 12, color: '#475569', display: 'flex', alignItems: 'center', gap: 4 }}>
                    <strong>Vessel:</strong>
                    {vesselVal ? (
                      <span style={{ fontWeight: 700, color: '#0f172a' }}>
                        {vesselVal}
                        <FieldConfidenceBadge confidence={vesselData.confidence} tier={vesselData.tier} />
                      </span>
                    ) : (
                      <span style={{ background: '#fee2e2', color: '#dc2626', border: '1px dashed #fca5a5', padding: '1px 6px', borderRadius: 4, fontWeight: 700, fontSize: 11 }}>
                        ⚠️ Not detected
                      </span>
                    )}
                  </div>
                  <div style={{ marginTop: 6, fontSize: 12, color: '#475569', display: 'flex', alignItems: 'center', gap: 4 }}>
                    <strong>🕒 Scanned:</strong>
                    <span>{formatScannedTime(item.updated_at || item.created_at)}</span>
                  </div>
                  <div style={{ marginTop: 6, fontSize: 12, color: '#475569' }}>
                    <strong>Source:</strong> {item.upload_source === 'folder' ? 'Folder' : 'Direct'}
                  </div>
                  <div style={{ marginTop: 10, display: 'grid', gridTemplateColumns: '1fr', gap: 8 }}>
                    <button type="button" onClick={() => handleOpenEdit(item)} style={{ minHeight: 44, background: clay.surfaceRaised, color: clay.text, border: `1px solid ${clay.accentSoft}`, borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Edit</button>
                    {item.status === 'needs_review' ? (
                      <button type="button" disabled={!!promotingIds[item.id]} onClick={() => handlePromoteToStaged(item)} style={{ minHeight: 44, background: promotingIds[item.id] ? '#94a3b8' : '#d97706', color: '#fff', border: 'none', borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: promotingIds[item.id] ? 'not-allowed' : 'pointer' }}>
                        {promotingIds[item.id] ? 'Moving...' : 'Move to Staging'}
                      </button>
                    ) : (
                      <button type="button" disabled={isMoving || isPending} onClick={() => handleMoveItem(item)} style={{ minHeight: 44, background: isMoving || isPending ? '#94a3b8' : '#15803d', color: '#fff', border: 'none', borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: isMoving || isPending ? 'not-allowed' : 'pointer' }}>
                        {isMoving ? 'Moving...' : 'Move'}
                      </button>
                    )}
                  </div>
                  <details style={{ marginTop: 10 }}>
                    <summary style={{ cursor: 'pointer', color: '#0369a1', fontSize: 12, fontWeight: 700 }}>Details</summary>
                    <div style={{ marginTop: 8, fontSize: 12, color: '#334155', lineHeight: 1.5 }}>
                      <div><strong>Group:</strong> {groupVal}</div>
                      <div><strong>Category:</strong> {catVal}</div>
                      <div><strong>Sub-Category:</strong> {subCatVal}</div>
                      <div><strong>Scanned Time:</strong> {formatScannedTime(item.updated_at || item.created_at)}</div>
                      <div style={{ wordBreak: 'break-all' }}><strong>Target Path:</strong> {item.final_path || `Technical & Crewing/${vesselVal || '{vessel}'}/Drawings and Manuals/${groupVal}/${catVal}`}</div>
                    </div>
                  </details>
                </div>
              );
            })}
          </div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', minWidth: isTablet ? 1120 : 980, borderCollapse: 'collapse', fontSize: 13, textAlign: 'left' }}>
            <thead>
              <tr style={{ background: clay.surfaceRaised, borderBottom: `1px solid ${clay.accentSoft}`, color: clay.textMuted, fontSize: 11, textTransform: 'uppercase' }}>
                {activeQueueTab === 'unstaged' && (
                  <th style={{ padding: '12px 12px', width: 36, textAlign: 'center' }}>
                    <input
                      type="checkbox"
                      checked={filteredUnstaged.length > 0 && selectedUnstagedIds.length === filteredUnstaged.length}
                      onChange={e => {
                        if (e.target.checked) {
                          setSelectedUnstagedIds(filteredUnstaged.map(i => i.id));
                        } else {
                          setSelectedUnstagedIds([]);
                        }
                      }}
                      title="Select all"
                      style={{ cursor: 'pointer' }}
                    />
                  </th>
                )}
                <th style={{ padding: '12px 16px', width: 56 }}>#</th>
                <th style={{ padding: '12px 16px', position: 'sticky', left: activeQueueTab === 'unstaged' ? 92 : 56, zIndex: 3, background: clay.surfaceRaised }}>Document Name</th>
                <th style={{ padding: '12px 16px' }}>Source</th>
                <th style={{ padding: '12px 16px' }}>Status &amp; AI Match</th>
                <th style={{ padding: '12px 16px' }}>SharePoint Metadata Tags</th>
                <th style={{ padding: '12px 16px', width: 290, minWidth: 290 }}>Target DMS Path</th>
                <th style={{ padding: '12px 16px', textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredItems.map((item, idx) => {
                const isMoving = !!movingIds[item.id];
                const isPending = item.status === 'ocr_pending';
                const confidence = item.confidence || 0;

                // Extract per-field data from suggested_tags
                const tags = item.suggested_tags || {};
                const catData = getFieldData(tags, 'category');
                const groupData = getFieldData(tags, 'group');
                const subCatData = getFieldData(tags, 'sub_category');
                const vesselData = getFieldData(tags, 'vessel');

                const catVal = catData.value || item.category_name || 'Drawing';
                const groupVal = groupData.value || '—';
                const subCatVal = subCatData.value || '—';
                // Strictly apply confidence floor to vessel value
                const vesselConfident = vesselData.confidence >= VESSEL_CONFIDENCE_FLOOR;
                const vesselVal = (vesselConfident ? (vesselData.value || item.vessel_name) : '') || '';

                const isSelected = selectedUnstagedIds.includes(item.id);

                // Check how many fields achieved high confidence (>= 85%)
                const confidentFieldsCount = [
                  catData.confidence >= 0.85,
                  groupData.confidence >= 0.85,
                  subCatData.confidence >= 0.85,
                  vesselData.confidence >= 0.85,
                ].filter(Boolean).length;

                return (
                  <tr key={item.id} style={{ borderBottom: `1px solid ${clay.accentSoft}`, background: isSelected ? clay.surfaceRaised : clay.surface }}>
                    {activeQueueTab === 'unstaged' && (
                      <td style={{ padding: '14px 12px', textAlign: 'center' }}>
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={e => {
                            if (e.target.checked) {
                              setSelectedUnstagedIds(prev => [...prev, item.id]);
                            } else {
                              setSelectedUnstagedIds(prev => prev.filter(id => id !== item.id));
                            }
                          }}
                          style={{ cursor: 'pointer' }}
                        />
                      </td>
                    )}
                    <td style={{ padding: '14px 16px', color: clay.textMuted, fontWeight: 600 }}>{idx + 1}</td>

                    {/* Filename, Preview & Scanned Time */}
                    <td style={{ padding: '14px 16px', maxWidth: 230, position: 'sticky', left: activeQueueTab === 'unstaged' ? 92 : 56, zIndex: 2, background: isSelected ? clay.surfaceRaised : clay.surface }}>
                      <div style={{ fontWeight: 700, color: clay.text, wordBreak: 'break-word' }}>
                        {item.filename}
                      </div>
                      {item.ocr_text_preview && (
                        <div style={{
                          fontSize: 11, color: '#64748b', marginTop: 3,
                          overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 210,
                        }}>
                          "{item.ocr_text_preview.slice(0, 60)}..."
                        </div>
                      )}
                      <div style={{
                        fontSize: 11, color: '#0369a1', marginTop: 5,
                        display: 'inline-flex', alignItems: 'center', gap: 4,
                        background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 4, padding: '2px 6px',
                        fontWeight: 600,
                      }}>
                        <span style={{ fontSize: 10 }}>🕒</span>
                        <span>Scanned: {formatScannedTime(item.updated_at || item.created_at)}</span>
                      </div>
                    </td>

                    {/* Source */}
                    <td style={{ padding: '14px 16px', whiteSpace: 'nowrap' }}>
                      <span style={{
                        padding: '3px 8px', borderRadius: 6, fontSize: 11, fontWeight: 700,
                        background: item.upload_source === 'folder' ? '#f0fdf4' : '#eff6ff',
                        color: item.upload_source === 'folder' ? '#15803d' : '#1d4ed8',
                        border: '1px solid',
                        borderColor: item.upload_source === 'folder' ? '#bbf7d0' : '#bfdbfe',
                      }}>
                        {item.upload_source === 'folder' ? '📁 Folder' : '📄 Direct'}
                      </span>
                    </td>

                    {/* Status & Confidence */}
                    <td style={{ padding: '14px 16px', whiteSpace: 'nowrap' }}>
                      {item.status === 'needs_review' ? (
                        <div>
                          <span style={{
                            padding: '3px 8px', borderRadius: 12, fontSize: 11, fontWeight: 700,
                            background: '#fee2e2', color: '#dc2626', border: '1px solid #fca5a5',
                            display: 'inline-block', marginBottom: 4,
                          }}>
                            ⚠️ Needs Review
                          </span>
                          <div style={{ fontSize: 10, color: '#64748b', marginTop: 2, display: 'flex', alignItems: 'center', gap: 3 }}>
                            <span>🕒 {formatScannedTime(item.updated_at || item.created_at)}</span>
                          </div>
                          {item.error && (
                            <div style={{ fontSize: 10, color: '#b91c1c', maxWidth: 180, whiteSpace: 'normal', lineHeight: 1.4, marginTop: 2 }}>
                              {item.error.replace(/^\s*\[/, '').replace(/\]\s*$/, '')}
                            </div>
                          )}
                        </div>
                      ) : isPending ? (
                        <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6, color: '#0284c7', fontWeight: 700, fontSize: 12 }}>
                          <span style={{ animation: 'spin 1.5s linear infinite' }}>⏳</span> OCR Extracting...
                        </div>
                      ) : (
                        <div>
                          <span style={{
                            padding: '3px 8px', borderRadius: 12, fontSize: 11, fontWeight: 700,
                            background: confidence >= 0.85 ? '#dcfce7' : confidence >= 0.60 ? '#fef3c7' : '#fee2e2',
                            color: confidence >= 0.85 ? '#15803d' : confidence >= 0.60 ? '#b45309' : '#dc2626',
                            border: '1px solid',
                            borderColor: confidence >= 0.85 ? '#86efac' : confidence >= 0.60 ? '#fde68a' : '#fca5a5',
                          }}>
                            {confidence >= 0.85 ? '🎯' : confidence >= 0.60 ? '⚡' : '⚠️'} {(confidence * 100).toFixed(0)}% Match ({confidentFieldsCount}/4 high)
                          </span>
                          <div style={{ fontSize: 10, color: '#64748b', marginTop: 3, display: 'flex', alignItems: 'center', gap: 3 }}>
                            <span>🕒 {formatScannedTime(item.updated_at || item.created_at)}</span>
                          </div>
                          {item.matched_keywords && item.matched_keywords.length > 0 && (
                            <div style={{ fontSize: 10, color: '#64748b', marginTop: 2 }}>
                              #{item.matched_keywords.slice(0, 2).join(' #')}
                            </div>
                          )}
                          {item.error && (
                            <div style={{ fontSize: 10, color: '#b45309', maxWidth: 180, whiteSpace: 'normal', lineHeight: 1.4, marginTop: 4 }}>
                              ⚠️ {item.error.replace(/^\s*\[/, '').replace(/\]\s*$/, '')}
                            </div>
                          )}
                        </div>
                      )}
                    </td>

                    {/* 4 SharePoint Columns with per-field confidence */}
                    <td style={{ padding: '14px 16px', minWidth: 260 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                        <span style={{ background: '#e0f2fe', color: '#0369a1', padding: '2px 6px', borderRadius: 4, fontSize: 11, fontWeight: 700 }}>
                          {groupVal}
                          <FieldConfidenceBadge confidence={groupData.confidence} tier={groupData.tier} />
                        </span>
                        <span style={{ color: '#94a3b8', fontSize: 11 }}>›</span>
                        <span style={{ background: '#f1f5f9', color: '#334155', padding: '2px 6px', borderRadius: 4, fontSize: 11, fontWeight: 600 }}>
                          {catVal}
                          <FieldConfidenceBadge confidence={catData.confidence} tier={catData.tier} />
                        </span>
                        <span style={{ color: '#94a3b8', fontSize: 11 }}>›</span>
                        <span style={{ color: '#0f172a', fontSize: 12, fontWeight: 700 }}>
                          {subCatVal}
                          <FieldConfidenceBadge confidence={subCatData.confidence} tier={subCatData.tier} />
                        </span>
                      </div>
                      {/* Vessel Row: shows vessel name + badge, or highlighted 'Not detected' pill */}
                      {vesselVal ? (
                        <div style={{ fontSize: 11, color: '#64748b', marginTop: 5, display: 'flex', alignItems: 'center' }}>
                          <span>🚢 Vessel:&nbsp;</span>
                          <strong style={{ color: '#0f172a' }}>{vesselVal}</strong>
                          <FieldConfidenceBadge confidence={vesselData.confidence} tier={vesselData.tier} />
                        </div>
                      ) : (
                        <div style={{ fontSize: 11, color: '#dc2626', marginTop: 5, display: 'flex', alignItems: 'center', gap: 4 }}>
                          <span>🚢 Vessel:&nbsp;</span>
                          <span style={{
                            background: '#fee2e2', color: '#dc2626', border: '1px dashed #fca5a5',
                            padding: '1px 6px', borderRadius: 4, fontWeight: 700, fontSize: 10,
                          }}>
                            ⚠️ Not detected
                          </span>
                          {vesselData.confidence > 0 && <FieldConfidenceBadge confidence={vesselData.confidence} tier={vesselData.tier} />}
                        </div>
                      )}
                    </td>

                    {/* Destination Path Preview */}
                    <td style={{ padding: '14px 16px', width: 290, minWidth: 290, verticalAlign: 'top' }}>
                      <div style={{
                        fontSize: 11, color: '#0f766e', fontFamily: 'monospace', fontWeight: 600,
                        background: '#f0fdfa', padding: '4px 8px', borderRadius: 6, border: '1px solid #ccfbf1',
                        lineHeight: 1.45,
                        whiteSpace: 'normal',
                        wordBreak: 'break-word',
                        overflowWrap: 'anywhere',
                      }}>
                        📁 {item.final_path || `Technical & Crewing/${vesselVal || '{vessel}'}/Drawings and Manuals/${groupVal}/${catVal}`}
                      </div>
                    </td>

                    {/* Actions */}
                    <td style={{ padding: '14px 16px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                      <div style={{ display: 'inline-flex', gap: 6, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                        <button
                          type="button"
                          onClick={() => handleOpenEdit(item)}
                          style={{
                            background: '#f1f5f9', color: '#334155', border: '1px solid #cbd5e1',
                            borderRadius: 6, padding: '5px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer',
                          }}
                        >
                          ✏️ Edit
                        </button>
                        {item.status === 'needs_review' ? (
                          <button
                            type="button"
                            disabled={!!promotingIds[item.id]}
                            onClick={() => handlePromoteToStaged(item)}
                            title="After assigning vessel and fields, click to move this file to the Staging Queue for filing"
                            style={{
                              background: promotingIds[item.id] ? '#94a3b8' : '#d97706',
                              color: '#fff', border: 'none', borderRadius: 6,
                              padding: '5px 12px', fontSize: 11, fontWeight: 700,
                              cursor: promotingIds[item.id] ? 'not-allowed' : 'pointer',
                            }}
                          >
                            {promotingIds[item.id] ? '⏳ Moving...' : '📋 Move to Staging'}
                          </button>
                        ) : (
                          <button
                            type="button"
                            disabled={isMoving || isPending}
                            onClick={() => handleMoveItem(item)}
                            style={{
                              background: isMoving || isPending ? '#94a3b8' : '#15803d',
                              color: '#fff', border: 'none', borderRadius: 6,
                              padding: '5px 12px', fontSize: 11, fontWeight: 700,
                              cursor: isMoving || isPending ? 'not-allowed' : 'pointer',
                            }}
                          >
                            {isMoving ? '⏳ Moving...' : '🚀 Move'}
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => handleDismissItem(item)}
                          style={{
                            background: 'none', color: '#94a3b8', border: 'none',
                            minHeight: 44, fontSize: 14, cursor: 'pointer', padding: '0 8px',
                          }}
                          title="Dismiss from queue"
                        >
                          ✕
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          </div>
        )}
      </div>

      {/* ── Edit Tags Modal ── */}
      {editingItem && (
        <div
          role="presentation"
          onClick={() => setEditingItem(null)}
          style={{
            position: 'fixed', inset: 0, zIndex: 10200,
            background: 'rgba(15,23,42,0.6)', backdropFilter: 'blur(4px)',
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 12 : 24,
          }}
        >
          <div
            role="dialog"
            aria-modal="true"
            onClick={e => e.stopPropagation()}
            style={{
              width: isMobile ? '95vw' : 600, maxWidth: '95vw', background: clay.surface, borderRadius: clay.radiusCard,
              boxShadow: clay.shadowRaisedHover, overflow: 'hidden',
              display: 'flex', flexDirection: 'column', maxHeight: '90vh',
            }}
          >
            {/* Modal Header */}
            <div style={{
              padding: '18px 24px', borderBottom: `1px solid ${clay.accentSoft}`, background: clay.surfaceRaised,
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            }}>
              <div>
                <div style={{ fontSize: 16, fontWeight: 800, color: clay.text }}>
                  Review &amp; Edit SharePoint Metadata Tags
                </div>
                <div style={{ fontSize: 12, color: clay.textMuted, marginTop: 2 }}>
                  {editingItem.filename}
                </div>
              </div>
              <button
                type="button"
                onClick={() => setEditingItem(null)}
                style={{ background: 'none', border: 'none', fontSize: 20, color: clay.textMuted, cursor: 'pointer' }}
              >
                ✕
              </button>
            </div>

            {/* Modal Body */}
            <div style={{ padding: '20px 24px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 }}>
              {editError && (
                <div style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#991b1b', padding: '10px 14px', borderRadius: 8, fontSize: 12 }}>
                  ❌ {editError}
                </div>
              )}

              {/* Secondary Multi-Engine Validation Section */}
              {(() => {
                const tags = editingItem.suggested_tags || {};
                const valInfo = (tags as any)._validation || {};
                const details = valInfo.field_details || {};
                const engines = valInfo.engines_used || ['primary_ocr'];
                const hasValidation = Object.keys(details).length > 0;

                if (!hasValidation) return null;

                return (
                  <div style={{
                    background: valInfo.overall_status === 'validated' ? '#f0fdf4' : '#fff7ed',
                    border: `1px solid ${valInfo.overall_status === 'validated' ? '#bbf7d0' : '#fed7aa'}`,
                    borderRadius: 10, padding: 14, display: 'flex', flexDirection: 'column', gap: 10,
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 6 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: valInfo.overall_status === 'validated' ? '#15803d' : '#c2410c', display: 'flex', alignItems: 'center', gap: 6 }}>
                        <span>{valInfo.overall_status === 'validated' ? '✅ Multi-Engine Cross-Validation: Validated' : '⚠️ Multi-Engine Cross-Validation: Review Suggested'}</span>
                      </div>
                      <span style={{ fontSize: 11, color: '#64748b' }}>Engines: {engines.join(', ')}</span>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, fontSize: 12 }}>
                      {['vessel', 'group', 'category', 'sub_category'].map(fKey => {
                        const fDetail = details[fKey];
                        if (!fDetail) return null;
                        const labelMap: Record<string, string> = { vessel: 'Vessel Name', group: 'Group', category: 'Category', sub_category: 'Sub-Category' };
                        const isMatch = fDetail.status === 'validated';
                        return (
                          <div key={fKey} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#fff', padding: '6px 10px', borderRadius: 6, border: '1px solid #e2e8f0', flexWrap: 'wrap', gap: 8 }}>
                            <div>
                              <strong style={{ color: '#334155' }}>{labelMap[fKey] || fKey}:</strong>{' '}
                              <span style={{ color: '#0f172a' }}>{fDetail.primary_value || '—'}</span>
                              {fDetail.secondary_value && fDetail.secondary_value !== fDetail.primary_value && (
                                <span style={{ color: '#c2410c', marginLeft: 8 }}>
                                  (Secondary: <strong>{fDetail.secondary_value}</strong>)
                                </span>
                              )}
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                              {fDetail.secondary_value && fDetail.secondary_value !== fDetail.primary_value && (
                                <button
                                  type="button"
                                  onClick={() => setEditTags(prev => ({ ...prev, [fKey]: fDetail.secondary_value }))}
                                  style={{
                                    background: '#f0f9ff', border: '1px solid #7dd3fc', borderRadius: 4,
                                    padding: '2px 8px', fontSize: 11, color: '#0369a1', cursor: 'pointer', fontWeight: 600,
                                  }}
                                >
                                  Use Secondary
                                </button>
                              )}
                              <span style={{ fontSize: 11, fontWeight: 700, color: isMatch ? '#15803d' : '#c2410c' }}>
                                {isMatch ? '✓ Agreed' : '⚠ Differing'}
                              </span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}

              {/* 4 SharePoint Managed Metadata Fields (Cascading) */}
              {(() => {
                const tags = editingItem.suggested_tags || {};
                const grpData = getFieldData(tags, 'group');
                const catData = getFieldData(tags, 'category');
                const subCatData = getFieldData(tags, 'sub_category');
                const vData = getFieldData(tags, 'vessel');

                return (
                  <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 14 }}>
                    {/* 1. Group (Drawing / Manual) */}
                    <div>
                      <label style={{ display: 'flex', alignItems: 'center', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                        📂 Group (Drawing / Manual) <span style={{ color: '#dc2626', marginLeft: 2 }}>*</span>
                        <FieldConfidenceBadge confidence={grpData.confidence} tier={grpData.tier} />
                      </label>
                      <select
                        value={editTags.group || 'Drawing'}
                        onChange={e => {
                          const newGroup = e.target.value;
                          const isManual = newGroup === 'Manual';
                          const catOptions = isManual ? Object.keys(MANUAL_TAXONOMY) : Object.keys(DRAWING_TAXONOMY);
                          const newCat = catOptions[0] || 'Basic';
                          const subCatOptions = (isManual ? MANUAL_TAXONOMY[newCat] : DRAWING_TAXONOMY[newCat]) || [];
                          setEditTags(prev => ({
                            ...prev,
                            group: newGroup,
                            category: newCat,
                            sub_category: subCatOptions[0] || '',
                          }));
                        }}
                        style={{
                          width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                          fontSize: 13, fontWeight: 600, color: '#0f172a', background: '#fff',
                        }}
                      >
                        <option value="Drawing">Drawing</option>
                        <option value="Manual">Manual</option>
                      </select>
                    </div>

                    {/* 2. Category (Cascading based on Group) */}
                    <div>
                      <label style={{ display: 'flex', alignItems: 'center', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                        📑 Category <span style={{ color: '#dc2626', marginLeft: 2 }}>*</span>
                        <FieldConfidenceBadge confidence={catData.confidence} tier={catData.tier} />
                      </label>
                      {(() => {
                        const isManual = editTags.group === 'Manual';
                        const catOptions = isManual ? Object.keys(MANUAL_TAXONOMY) : Object.keys(DRAWING_TAXONOMY);
                        return (
                          <select
                            value={editTags.category || catOptions[0]}
                            onChange={e => {
                              const newCat = e.target.value;
                              const availableSubCats = (isManual ? MANUAL_TAXONOMY[newCat] : DRAWING_TAXONOMY[newCat]) || [];
                              setEditTags(prev => ({
                                ...prev,
                                category: newCat,
                                sub_category: availableSubCats[0] || '',
                              }));
                            }}
                            style={{
                              width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                              fontSize: 13, fontWeight: 600, color: '#0f172a', background: '#fff',
                            }}
                          >
                            {catOptions.map(c => (
                              <option key={c} value={c}>{c}</option>
                            ))}
                          </select>
                        );
                      })()}
                    </div>

                    {/* 3. Sub-Category (Cascading based on Category) */}
                    <div style={{ gridColumn: isMobile ? 'span 1' : 'span 2' }}>
                      <label style={{ display: 'flex', alignItems: 'center', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                        🏷️ Sub-Category (Specific Term) <span style={{ color: '#dc2626', marginLeft: 2 }}>*</span>
                        <FieldConfidenceBadge confidence={subCatData.confidence} tier={subCatData.tier} />
                      </label>
                      {(() => {
                        const isManual = editTags.group === 'Manual';
                        const currentCat = editTags.category || (isManual ? Object.keys(MANUAL_TAXONOMY)[0] : Object.keys(DRAWING_TAXONOMY)[0]);
                        const subCatOptions = (isManual ? MANUAL_TAXONOMY[currentCat] : DRAWING_TAXONOMY[currentCat]) || [];
                        return (
                          <select
                            value={editTags.sub_category || (subCatOptions[0] || '')}
                            onChange={e => setEditTags(prev => ({ ...prev, sub_category: e.target.value }))}
                            style={{
                              width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                              fontSize: 13, fontWeight: 600, color: '#0f172a', background: '#fff',
                            }}
                          >
                            {subCatOptions.map(sc => (
                              <option key={sc} value={sc}>{sc}</option>
                            ))}
                          </select>
                        );
                      })()}
                    </div>

                    {/* 4. Vessel Name */}
                    <div style={{ gridColumn: isMobile ? 'span 1' : 'span 2' }}>
                      <label style={{ display: 'flex', alignItems: 'center', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                        🚢 Vessel Name <span style={{ color: '#dc2626', marginLeft: 2 }}>*</span>
                        <FieldConfidenceBadge confidence={vData.confidence} tier={vData.tier} />
                      </label>
                      <select
                        value={editTags.vessel || ''}
                        onChange={e => setEditTags(prev => ({ ...prev, vessel: e.target.value }))}
                        style={{
                          width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                          fontSize: 13, fontWeight: 600, color: '#0f172a', background: '#fff',
                        }}
                      >
                        <option value="">-- Select Vessel --</option>
                        {vesselOptions.map(v => (
                          <option key={v} value={v}>{v}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                );
              })()}

              {/* Dynamic SharePoint Destination Path Preview */}
              <div style={{
                background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: '12px 16px',
              }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>
                  Live Destination Path Preview
                </div>
                <div style={{
                  fontSize: 12, color: '#0f766e', fontFamily: 'monospace', fontWeight: 600,
                  marginTop: 4, wordBreak: 'break-all',
                }}>
                  📁 {livePreviewPath || `Technical & Crewing/${editTags.vessel || '[Vessel]'}/Drawings and Manuals/${editTags.group || 'Drawing'}/${editTags.category || 'Basic'}`}
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div style={{
              padding: '14px 24px', borderTop: `1px solid ${clay.accentSoft}`, background: clay.surfaceRaised,
              display: 'flex', justifyContent: 'flex-end', gap: 10, flexDirection: isMobile ? 'column' : 'row',
            }}>
              <button
                type="button"
                onClick={() => setEditingItem(null)}
                style={{
                  background: clay.surface, color: clay.textMuted, border: `1px solid ${clay.accentSoft}`,
                  borderRadius: 8, padding: '8px 16px', minHeight: 44, width: isMobile ? '100%' : 'auto', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={isSavingEdit}
                onClick={handleSaveTags}
                style={{
                  background: clay.accentGradient, color: '#fff', border: 'none',
                  borderRadius: 8, padding: '8px 20px', minHeight: 44, width: isMobile ? '100%' : 'auto', fontSize: 13, fontWeight: 700,
                  cursor: isSavingEdit ? 'wait' : 'pointer',
                  boxShadow: clay.shadowButton,
                }}
              >
                {isSavingEdit ? '⏳ Saving...' : '💾 Save Tags'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── "Which Vessel Is This?" Single-File Action Prompt Modal (Option A) ── */}
      {vesselPromptItem && (
        <div
          role="presentation"
          onClick={() => setVesselPromptItem(null)}
          style={{
            position: 'fixed', inset: 0, zIndex: 10300,
            background: 'rgba(15,23,42,0.65)', backdropFilter: 'blur(4px)',
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 12 : 24,
          }}
        >
          <div
            role="dialog"
            aria-modal="true"
            onClick={e => e.stopPropagation()}
            style={{
              width: isMobile ? '95vw' : 520, maxWidth: '95vw', background: '#fff', borderRadius: 16,
              boxShadow: '0 25px 60px rgba(0,0,0,0.3)', overflow: 'hidden',
              display: 'flex', flexDirection: 'column',
            }}
          >
            {/* Header */}
            <div style={{
              padding: '18px 24px', borderBottom: `1px solid ${clay.accentSoft}`, background: 'linear-gradient(135deg, #edfdf9 0%, #dff7ff 100%)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start',
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <div style={{
                  width: 40, height: 40, borderRadius: 10, background: '#0284c7', color: '#fff',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, flexShrink: 0,
                }}>
                  🚢
                </div>
                <div>
                  <div style={{ fontSize: 16, fontWeight: 800, color: '#0f172a' }}>
                    Which vessel is this?
                  </div>
                  <div style={{ fontSize: 12, color: '#0369a1', marginTop: 2, fontWeight: 600 }}>
                    Vessel required before {vesselPromptItem.action === 'promote' ? 'moving to staging' : 'moving to SharePoint'}
                  </div>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setVesselPromptItem(null)}
                style={{ background: 'none', border: 'none', fontSize: 20, color: '#64748b', cursor: 'pointer' }}
              >
                ✕
              </button>
            </div>

            {/* Body */}
            <div style={{ padding: '20px 24px', display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div style={{
                background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: '10px 14px',
                fontSize: 12, color: '#334155', display: 'flex', alignItems: 'center', gap: 8,
              }}>
                <span style={{ fontSize: 16 }}>📄</span>
                <strong style={{ wordBreak: 'break-all' }}>{vesselPromptItem.item.filename}</strong>
              </div>

              <div style={{ fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
                The vessel could not be identified automatically with high confidence. Please select which vessel this document belongs to:
              </div>

              {promptError && (
                <div style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#991b1b', padding: '8px 12px', borderRadius: 8, fontSize: 12 }}>
                  ❌ {promptError}
                </div>
              )}

              {!isCreatingNewVesselInline ? (
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 6 }}>
                    Select Vessel <span style={{ color: '#dc2626' }}>*</span>
                  </label>
                  <select
                    value={promptVesselChoice}
                    onChange={e => setPromptVesselChoice(e.target.value)}
                    style={{
                      width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1',
                      fontSize: 13, fontWeight: 600, color: '#0f172a', background: '#fff', outline: 'none',
                    }}
                  >
                    <option value="">-- Choose a vessel --</option>
                    {vesselOptions.map(v => (
                      <option key={v} value={v}>{v}</option>
                    ))}
                  </select>

                  <div style={{ marginTop: 8, textAlign: 'right' }}>
                    <button
                      type="button"
                      onClick={() => setIsCreatingNewVesselInline(true)}
                      style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
                    >
                      + Create new vessel not in list
                    </button>
                  </div>
                </div>
              ) : (
                <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 10, padding: 12 }}>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#0369a1', marginBottom: 6 }}>
                    New Vessel Name <span style={{ color: '#dc2626' }}>*</span>
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. Pacific Explorer"
                    value={newVesselInput}
                    onChange={e => setNewVesselInput(e.target.value)}
                    style={{
                      width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #93c5fd',
                      fontSize: 13, fontWeight: 600, color: '#0f172a', outline: 'none',
                    }}
                  />
                  <div style={{ marginTop: 8, textAlign: 'right' }}>
                    <button
                      type="button"
                      onClick={() => setIsCreatingNewVesselInline(false)}
                      style={{ background: 'none', border: 'none', color: '#64748b', fontSize: 12, cursor: 'pointer' }}
                    >
                      ← Back to existing vessels list
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Footer */}
            <div style={{
              padding: '14px 24px', borderTop: `1px solid ${clay.accentSoft}`, background: clay.surfaceRaised,
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10,
            }}>
              <button
                type="button"
                onClick={() => setVesselPromptItem(null)}
                style={{
                  background: '#fff', color: '#64748b', border: '1px solid #cbd5e1',
                  borderRadius: 8, padding: '8px 16px', minHeight: 40, fontSize: 13, fontWeight: 600, cursor: 'pointer',
                }}
              >
                Skip for Now
              </button>

              <button
                type="button"
                disabled={isSavingPrompt}
                onClick={handleConfirmVesselPrompt}
                style={{
                  background: '#0284c7', color: '#fff', border: 'none',
                  borderRadius: 8, padding: '8px 20px', minHeight: 40, fontSize: 13, fontWeight: 700,
                  cursor: isSavingPrompt ? 'wait' : 'pointer',
                  boxShadow: '0 2px 6px rgba(2,132,199,0.3)',
                }}
              >
                {isSavingPrompt ? '⏳ Saving & Proceeding...' : vesselPromptItem.action === 'promote' ? 'Confirm & Move to Staging' : 'Confirm & Move'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Batch Vessel Assignment Modal (Option 4b) ── */}
      {batchModalOpen && (
        <div
          role="presentation"
          onClick={() => setBatchModalOpen(false)}
          style={{
            position: 'fixed', inset: 0, zIndex: 10300,
            background: 'rgba(15,23,42,0.65)', backdropFilter: 'blur(4px)',
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 12 : 24,
          }}
        >
          <div
            role="dialog"
            aria-modal="true"
            onClick={e => e.stopPropagation()}
            style={{
              width: isMobile ? '95vw' : 520, maxWidth: '95vw', background: '#fff', borderRadius: 16,
              boxShadow: '0 25px 60px rgba(0,0,0,0.3)', overflow: 'hidden',
              display: 'flex', flexDirection: 'column',
            }}
          >
            {/* Header */}
            <div style={{
              padding: '18px 24px', borderBottom: '1px solid #e2e8f0', background: '#f8fafc',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            }}>
              <div>
                <div style={{ fontSize: 16, fontWeight: 800, color: '#0f172a' }}>
                  🚢 Batch Assign Vessel
                </div>
                <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>
                  Assign a single vessel to {selectedUnstagedIds.length} selected document{selectedUnstagedIds.length > 1 ? 's' : ''}
                </div>
              </div>
              <button
                type="button"
                onClick={() => setBatchModalOpen(false)}
                style={{ background: 'none', border: 'none', fontSize: 20, color: '#94a3b8', cursor: 'pointer' }}
              >
                ✕
              </button>
            </div>

            {/* Body */}
            <div style={{ padding: '20px 24px', display: 'flex', flexDirection: 'column', gap: 14 }}>
              {batchError && (
                <div style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#991b1b', padding: '8px 12px', borderRadius: 8, fontSize: 12 }}>
                  ❌ {batchError}
                </div>
              )}

              <div>
                <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 6 }}>
                  Target Vessel <span style={{ color: '#dc2626' }}>*</span>
                </label>
                <select
                  value={batchVesselChoice}
                  onChange={e => setBatchVesselChoice(e.target.value)}
                  style={{
                    width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1',
                    fontSize: 13, fontWeight: 600, color: '#0f172a', background: '#fff', outline: 'none',
                  }}
                >
                  <option value="">-- Select vessel to apply to all selected --</option>
                  {vesselOptions.map(v => (
                    <option key={v} value={v}>{v}</option>
                  ))}
                </select>
              </div>

              <div style={{
                background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '12px 14px',
                display: 'flex', alignItems: 'center', gap: 10,
              }}>
                <input
                  type="checkbox"
                  id="batchPromoteCheckbox"
                  checked={batchPromoteImmediately}
                  onChange={e => setBatchPromoteImmediately(e.target.checked)}
                  style={{ cursor: 'pointer', width: 16, height: 16 }}
                />
                <label htmlFor="batchPromoteCheckbox" style={{ fontSize: 12, fontWeight: 600, color: '#334155', cursor: 'pointer' }}>
                  Also promote all selected files directly to Staging Queue
                </label>
              </div>
            </div>

            {/* Footer */}
            <div style={{
              padding: '14px 24px', borderTop: '1px solid #e2e8f0', background: '#f8fafc',
              display: 'flex', justifyContent: 'flex-end', gap: 10,
            }}>
              <button
                type="button"
                onClick={() => setBatchModalOpen(false)}
                style={{
                  background: '#fff', color: '#64748b', border: '1px solid #cbd5e1',
                  borderRadius: 8, padding: '8px 16px', minHeight: 40, fontSize: 13, fontWeight: 600, cursor: 'pointer',
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={isApplyingBatch}
                onClick={handleApplyBatchVessel}
                style={{
                  background: '#0284c7', color: '#fff', border: 'none',
                  borderRadius: 8, padding: '8px 20px', minHeight: 40, fontSize: 13, fontWeight: 700,
                  cursor: isApplyingBatch ? 'wait' : 'pointer',
                  boxShadow: '0 2px 6px rgba(2,132,199,0.3)',
                }}
              >
                {isApplyingBatch ? '⏳ Applying...' : `Apply to ${selectedUnstagedIds.length} File${selectedUnstagedIds.length > 1 ? 's' : ''}`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

