/* eslint-disable @typescript-eslint/no-unused-vars */
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge } from '../constants';
import type { DocumentCategory, TagFieldDef } from '../types/ui';
import { OcrStagingQueue, formatScannedTime } from './OcrStagingQueue';
import { isMobileWidth, isTabletWidth } from '../responsive';

interface OcrClassificationResult {
  filename: string;
  file_size: number;
  content_type: string;
  text_preview: string;
  text_length: number;
  detected_vessel: string | null;
  detected_group: string;
  detected_category: string;
  detected_sub_category?: string;
  detected_sub_category_1: string;
  detected_sub_category_2: string;
  detected_leaf: string;
  suggested_path: string;
  confidence: number;
  matched_keywords: string[];
  matched_category_id?: number | null;
  matched_category_name?: string | null;
  tag_fields?: TagFieldDef[];
  suggested_tags?: Record<string, any>;
  available_vessels: string[];
  available_departments: string[];
  drawing_taxonomy: Record<string, string[]>;
  manual_taxonomy: Record<string, string[]>;
}

export function renderTemplatesPage(host: VesselEmail): React.ReactElement {
  return <TemplatesPageContainer host={host} />;
}

interface TemplatesPageContainerProps {
  host: VesselEmail;
}

const TemplatesPageContainer: React.FC<TemplatesPageContainerProps> = ({ host }) => {
  const { vessels } = host.state;
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const isMobile = isMobileWidth(viewportWidth);
  const isTablet = isTabletWidth(viewportWidth);

  // Active Top Tab: 'queue' | 'ocr' | 'categories'
  const [activeTab, setActiveTab] = React.useState<'queue' | 'ocr' | 'categories'>('queue');

  // Categories list from API
  const [categories, setCategories] = React.useState<DocumentCategory[]>([]);
  const [loadingCategories, setLoadingCategories] = React.useState<boolean>(true);

  // Category Editor Modal state
  const [editingCategory, setEditingCategory] = React.useState<DocumentCategory | null>(null);
  const [isNewCategoryModal, setIsNewCategoryModal] = React.useState<boolean>(false);
  const [catName, setCatName] = React.useState<string>('');
  const [catDept, setCatDept] = React.useState<string>('Technical & Crewing');
  const [catPathTemplate, setCatPathTemplate] = React.useState<string>('{group}/{vessel}/Drawings and Manuals/{category}/{sub_category}');
  const [catTagFields, setCatTagFields] = React.useState<TagFieldDef[]>([
    { key: 'vessel', label: 'Vessel Name', type: 'select_vessel', required: true, options: null },
    { key: 'group', label: 'Department', type: 'select_dept', required: true, options: null },
    { key: 'category', label: 'Category', type: 'text', required: true, options: null },
    { key: 'sub_category', label: 'Sub-Category', type: 'text', required: false, options: null },
  ]);
  const [catOcrHintsText, setCatOcrHintsText] = React.useState<string>('');
  const [isSavingCategory, setIsSavingCategory] = React.useState<boolean>(false);
  const [categoryModalError, setCategoryModalError] = React.useState<string | null>(null);

  // OCR Module Local State
  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [isExtracting, setIsExtracting] = React.useState<boolean>(false);
  const [extractError, setExtractError] = React.useState<string | null>(null);
  const [ocrResult, setOcrResult] = React.useState<OcrClassificationResult | null>(null);
  const [extractedAt, setExtractedAt] = React.useState<string | null>(null);

  // Editable dynamic fields after extraction
  const [dynamicTags, setDynamicTags] = React.useState<Record<string, any>>({});
  const [selectedMatchCatId, setSelectedMatchCatId] = React.useState<number | null>(null);

  // Upload & routing state
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
  const [uploadSuccess, setUploadSuccess] = React.useState<{ path: string; filename: string } | null>(null);
  const [uploadError, setUploadError] = React.useState<string | null>(null);

  // Existing file mode (triggered from List/Folder View via ocrPendingItemId)
  const [existingMode, setExistingMode] = React.useState<{ itemId: string; filename: string } | null>(null);
  const [isMoveTagging, setIsMoveTagging] = React.useState<boolean>(false);
  const [moveTagSuccess, setMoveTagSuccess] = React.useState<{ path: string; filename: string } | null>(null);

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  // Fetch categories from API
  const fetchCategories = React.useCallback(async () => {
    setLoadingCategories(true);
    try {
      const base = host._base();
      const res = await fetch(`${base}/api/categories`, { headers: host._headers() });
      if (res.ok) {
        const data: DocumentCategory[] = await res.json();
        setCategories(data);
      }
    } catch {
      // Fallback
    } finally {
      setLoadingCategories(false);
    }
  }, [host]);

  React.useEffect(() => {
    void fetchCategories();
  }, [fetchCategories]);

  // ── Detect ocrPendingItemId from host state ──
  React.useEffect(() => {
    const { ocrPendingItemId, ocrPendingFilename } = host.state;
    if (!ocrPendingItemId) return;

    host.setState({ ocrPendingItemId: null, ocrPendingFilename: null });

    setActiveTab('ocr');
    setExistingMode({ itemId: ocrPendingItemId, filename: ocrPendingFilename || 'document' });
    setOcrResult(null);
    setSelectedFile(null);
    setUploadSuccess(null);
    setMoveTagSuccess(null);
    setExtractError(null);
    setIsExtracting(true);

    const base = host._base();
    const form = new FormData();
    form.append('item_id', ocrPendingItemId);
    form.append('filename', ocrPendingFilename || 'document');

    fetch(`${base}/api/ocr/classify-existing-file`, {
      method: 'POST',
      headers: host._uploadHeaders(),
      body: form,
    })
      .then(async res => {
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || `Server returned status ${res.status}`);
        }
        return res.json() as Promise<OcrClassificationResult>;
      })
      .then(data => {
        setOcrResult(data);
        setSelectedMatchCatId(data.matched_category_id || null);
        setDynamicTags({
          vessel: data.suggested_tags?.vessel || data.detected_vessel || '',
          group: data.suggested_tags?.group || data.detected_group || 'Drawing',
          category: data.suggested_tags?.category || data.detected_category || 'Basic',
          sub_category: data.suggested_tags?.sub_category || data.detected_sub_category || '',
          department: data.suggested_tags?.department || 'Technical & Crewing',
          ...(data.suggested_tags || {}),
        });
      })
      .catch((err: any) => {
        setExtractError(err?.message || 'Failed to classify existing file.');
      })
      .finally(() => setIsExtracting(false));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [host.state.ocrPendingItemId]);

  const vesselOptions = React.useMemo(() => {
    const list = (vessels || []).map(v => v.name).filter(Boolean);
    if (ocrResult?.available_vessels) {
      ocrResult.available_vessels.forEach(v => {
        if (!list.includes(v)) list.push(v);
      });
    }
    return list.sort();
  }, [vessels, ocrResult]);

  // Active Category in OCR manual tab
  const activeOcrCategory = React.useMemo(() => {
    if (selectedMatchCatId) {
      const found = categories.find(c => c.id === selectedMatchCatId);
      if (found) return found;
    }
    return categories.find(c => c.name.toLowerCase() === (dynamicTags.category || '').toLowerCase()) || categories[0] || null;
  }, [categories, selectedMatchCatId, dynamicTags.category]);

  // Destination path resolved in OCR manual tab (folder ends at Category level)
  const resolvedOcrPath = React.useMemo(() => {
    let templateStr = activeOcrCategory?.dms_path_template || 'Technical & Crewing/{vessel}/Drawings and Manuals/{group}/{category}';
    templateStr = templateStr.replace(/\/\{sub_?category\}/gi, '');
    let p = templateStr;
    for (const [k, v] of Object.entries(dynamicTags)) {
      p = p.replace(new RegExp(`\\{${k}\\}`, 'g'), String(v || '').trim());
    }
    p = p.replace(/\{[a-z0-9_]+\}/g, '');
    p = p.replace(/\/+/g, '/').replace(/^\/+|\/+$/g, '');
    return p || 'Technical & Crewing/All Vessels/Drawings and Manuals';
  }, [activeOcrCategory, dynamicTags]);

  // Handle manual file selection & extraction
  const handleFileChange = async (file: File | undefined) => {
    if (!file) return;
    setSelectedFile(file);
    setIsExtracting(true);
    setExtractError(null);
    setOcrResult(null);
    setUploadSuccess(null);
    setUploadError(null);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const base = host._base();
      const response = await fetch(`${base}/api/ocr/classify`, {
        method: 'POST',
        headers: host._uploadHeaders(),
        body: formData,
      });

      if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        throw new Error(errJson.detail || `Server returned status ${response.status}`);
      }

      const data: OcrClassificationResult = await response.json();
      setOcrResult(data);
      setExtractedAt(new Date().toISOString());
      setSelectedMatchCatId(data.matched_category_id || null);

      setDynamicTags({
        vessel: data.suggested_tags?.vessel || data.detected_vessel || '',
        group: data.suggested_tags?.group || data.detected_group || 'Drawing',
        category: data.suggested_tags?.category || data.detected_category || 'Basic',
        sub_category: data.suggested_tags?.sub_category || data.detected_sub_category || '',
        department: data.suggested_tags?.department || 'Technical & Crewing',
        ...(data.suggested_tags || {}),
      });
    } catch (err: any) {
      setExtractError(err?.message || 'Failed to extract text and classify document.');
    } finally {
      setIsExtracting(false);
    }
  };

  // Move & Tag for existing file
  const handleMoveAndTag = async () => {
    if (!existingMode) return;
    setIsMoveTagging(true);
    setUploadError(null);
    setMoveTagSuccess(null);

    try {
      const base = host._base();
      const form = new FormData();
      form.append('item_id', existingMode.itemId);
      form.append('target_path', resolvedOcrPath);
      form.append('vessel_name', dynamicTags.vessel || '');
      form.append('group', dynamicTags.group || 'Drawing');
      form.append('category', dynamicTags.category || 'To be Classified');
      form.append('sub_category', dynamicTags.sub_category || 'To be Classified');
      form.append('department', dynamicTags.department || 'Technical & Crewing');

      const res = await fetch(`${base}/api/ocr/move-and-tag-existing`, {
        method: 'POST',
        headers: host._uploadHeaders(),
        body: form,
      });
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.detail || `Move failed with status ${res.status}`);
      }
      setMoveTagSuccess({ path: resolvedOcrPath, filename: existingMode.filename });
      setExistingMode(null);
      if (dynamicTags.vessel) {
        void host._mergeLiveSharePointFiles([dynamicTags.vessel]).catch(() => undefined);
      }
    } catch (err: any) {
      setUploadError(err?.message || 'Failed to move and tag document.');
    } finally {
      setIsMoveTagging(false);
    }
  };

  // Direct route & upload
  const handleRouteAndUpload = async () => {
    if (!selectedFile) return;
    setIsUploading(true);
    setUploadError(null);
    setUploadSuccess(null);

    try {
      const base = host._base();
      const formData = new FormData();
      formData.append('file', selectedFile);
      formData.append('path', resolvedOcrPath);
      formData.append('vessel_name', dynamicTags.vessel || '');
      formData.append('group', dynamicTags.group || 'Technical & Crewing');
      formData.append('category', dynamicTags.category || 'Drawings and Manuals');
      formData.append('sub_category', dynamicTags.sub_category || '');

      const response = await fetch(`${base}/api/ocr/route-and-upload`, {
        method: 'POST',
        headers: host._uploadHeaders(),
        body: formData,
      });

      if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        throw new Error(errJson.detail || `Upload failed with status ${response.status}`);
      }

      setUploadSuccess({ path: resolvedOcrPath, filename: selectedFile.name });
      if (dynamicTags.vessel) {
        void host._mergeLiveSharePointFiles([dynamicTags.vessel]).catch(() => undefined);
      }
    } catch (err: any) {
      setUploadError(err?.message || 'Failed to route and upload document to SharePoint.');
    } finally {
      setIsUploading(false);
    }
  };

  // ── Open Category Create Modal ──
  const handleOpenNewCategory = () => {
    setEditingCategory(null);
    setCatName('');
    setCatDept('Technical & Crewing');
    setCatPathTemplate('{group}/{vessel}/Drawings and Manuals/{category}/{sub_category}');
    setCatTagFields([
      { key: 'vessel', label: 'Vessel Name', type: 'select_vessel', required: true, options: null },
      { key: 'group', label: 'Department', type: 'select_dept', required: true, options: null },
      { key: 'category', label: 'Category', type: 'text', required: true, options: null },
      { key: 'sub_category', label: 'Sub-Category', type: 'text', required: false, options: null },
    ]);
    setCatOcrHintsText('');
    setCategoryModalError(null);
    setIsNewCategoryModal(true);
  };

  // ── Open Category Edit Modal ──
  const handleOpenEditCategory = (cat: DocumentCategory) => {
    setEditingCategory(cat);
    setCatName(cat.name);
    setCatDept(cat.department || 'Technical & Crewing');
    setCatPathTemplate(cat.dms_path_template || '{group}/{vessel}/Drawings and Manuals/{category}/{sub_category}');
    setCatTagFields(cat.tag_fields && cat.tag_fields.length > 0 ? [...cat.tag_fields] : [
      { key: 'vessel', label: 'Vessel Name', type: 'select_vessel', required: true, options: null },
      { key: 'group', label: 'Department', type: 'select_dept', required: true, options: null },
      { key: 'category', label: 'Category', type: 'text', required: true, options: null },
      { key: 'sub_category', label: 'Sub-Category', type: 'text', required: false, options: null },
    ]);
    setCatOcrHintsText((cat.ocr_hints || []).join(', '));
    setCategoryModalError(null);
    setIsNewCategoryModal(true);
  };

  // ── Save Category (Create or Update) ──
  const handleSaveCategory = async () => {
    if (!catName.trim()) {
      setCategoryModalError('Category name is required.');
      return;
    }

    // Validate unique keys in tag fields
    const keys = catTagFields.map(f => f.key.trim().toLowerCase());
    if (new Set(keys).size !== keys.length) {
      setCategoryModalError('All Tag Field Keys must be unique.');
      return;
    }

    const hints = catOcrHintsText
      .split(/[,;\n]/)
      .map(s => s.trim().toLowerCase())
      .filter(Boolean);

    setIsSavingCategory(true);
    setCategoryModalError(null);
    try {
      const base = host._base();
      const payload = {
        name: catName.trim(),
        department: catDept,
        dms_path_template: catPathTemplate.trim(),
        tag_fields: catTagFields,
        ocr_hints: hints,
      };

      let res: Response;
      if (editingCategory) {
        res = await fetch(`${base}/api/categories/${editingCategory.id}`, {
          method: 'PATCH',
          headers: host._headers(),
          body: JSON.stringify(payload),
        });
      } else {
        res = await fetch(`${base}/api/categories`, {
          method: 'POST',
          headers: host._headers(),
          body: JSON.stringify(payload),
        });
      }

      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.detail || `Failed to save category (status ${res.status})`);
      }

      setIsNewCategoryModal(false);
      void fetchCategories();
    } catch (err: any) {
      setCategoryModalError(err?.message || 'Error saving category');
    } finally {
      setIsSavingCategory(false);
    }
  };

  // ── Delete Category ──
  const handleDeleteCategory = async (cat: DocumentCategory) => {
    if (!window.confirm(`Deactivate document category "${cat.name}"?`)) return;
    try {
      const base = host._base();
      await fetch(`${base}/api/categories/${cat.id}`, {
        method: 'DELETE',
        headers: host._headers(),
      });
      void fetchCategories();
    } catch (err: any) {
      alert(`Delete error: ${err?.message}`);
    }
  };

  const stagingBadgeCount = host.state.ocrStagingCount || 0;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* Header & Tabs */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
            <span>📑</span> Templates, Categories &amp; OCR Staging
          </h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
            Manage category taxonomies, review OCR-staged documents, and auto-classify vessel files.
          </p>
        </div>

        {/* 3-Tab Pill Navigation */}
        <div style={{
          display: 'inline-flex', background: '#f1f5f9', borderRadius: 12, padding: 4,
          border: '1px solid #e2e8f0', boxShadow: 'inset 0 1px 2px rgba(0,0,0,0.04)',
          flexWrap: 'wrap',
        }}>
          {/* Tab 1: Queue */}
          <button
            type="button"
            onClick={() => setActiveTab('queue')}
            style={{
              padding: '8px 18px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              background: activeTab === 'queue' ? '#0284c7' : 'transparent',
              color: activeTab === 'queue' ? '#fff' : '#475569',
              boxShadow: activeTab === 'queue' ? '0 2px 6px rgba(2, 132, 199, 0.35)' : 'none',
              display: 'inline-flex', alignItems: 'center', gap: 8, transition: 'all 0.15s ease',
            }}
          >
            <span>📥</span> Review Staging Queue
            {stagingBadgeCount > 0 && (
              <span style={{
                fontSize: 11, background: activeTab === 'queue' ? '#fff' : '#0284c7',
                color: activeTab === 'queue' ? '#0284c7' : '#fff', padding: '1px 7px', borderRadius: 12, fontWeight: 800,
              }}>
                {stagingBadgeCount}
              </span>
            )}
          </button>

          {/* Tab 2: OCR Extractor */}
          <button
            type="button"
            onClick={() => setActiveTab('ocr')}
            style={{
              padding: '8px 18px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              background: activeTab === 'ocr' ? '#0284c7' : 'transparent',
              color: activeTab === 'ocr' ? '#fff' : '#475569',
              boxShadow: activeTab === 'ocr' ? '0 2px 6px rgba(2, 132, 199, 0.35)' : 'none',
              display: 'inline-flex', alignItems: 'center', gap: 8, transition: 'all 0.15s ease',
            }}
          >
            <span>🔍</span> Manual OCR Extractor
          </button>

          {/* Tab 3: Categories & Templates */}
          <button
            type="button"
            onClick={() => setActiveTab('categories')}
            style={{
              padding: '8px 18px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              background: activeTab === 'categories' ? '#0284c7' : 'transparent',
              color: activeTab === 'categories' ? '#fff' : '#475569',
              boxShadow: activeTab === 'categories' ? '0 2px 6px rgba(2, 132, 199, 0.35)' : 'none',
              display: 'inline-flex', alignItems: 'center', gap: 8, transition: 'all 0.15s ease',
            }}
          >
            <span>⚙️</span> Document Categories &amp; Tags
          </button>
        </div>
      </div>

      {/* ── TAB 1: REVIEW STAGING QUEUE ── */}
      {activeTab === 'queue' && (
        <OcrStagingQueue
          host={host}
          categories={categories}
          onRefreshCategories={fetchCategories}
        />
      )}

      {/* ── TAB 2: MANUAL OCR EXTRACTOR ── */}
      {activeTab === 'ocr' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
          {/* Drop Zone */}
          <div
            onDragOver={e => e.preventDefault()}
            onDrop={e => {
              e.preventDefault();
              if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                void handleFileChange(e.dataTransfer.files[0]);
              }
            }}
            style={{
              background: '#ffffff', borderRadius: 14, border: '2px dashed #93c5fd',
              padding: '36px 24px', textAlign: 'center', display: 'flex', flexDirection: 'column',
              alignItems: 'center', gap: 12, cursor: 'pointer',
            }}
            onClick={() => fileInputRef.current?.click()}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".pdf,.png,.jpg,.jpeg,.tiff,.bmp,.docx,.xlsx,.txt"
              style={{ display: 'none' }}
              onChange={e => {
                if (e.target.files && e.target.files[0]) {
                  void handleFileChange(e.target.files[0]);
                }
              }}
            />
            <div style={{
              width: 56, height: 56, borderRadius: '50%', background: '#eff6ff',
              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 26, color: '#0284c7',
            }}>
              📤
            </div>
            <div>
              <div style={{ fontSize: 16, fontWeight: 700, color: '#1e293b' }}>
                Upload Drawing or Document for Live OCR Classification
              </div>
              <div style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>
                Supports PDF, Scanned Images (PNG, JPG, TIFF), Word (.docx), and Technical Drawings
              </div>
            </div>
            <button
              type="button"
              style={{
                background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
                padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
              onClick={e => {
                e.stopPropagation();
                fileInputRef.current?.click();
              }}
            >
              📁 Browse Local File
            </button>
          </div>

          {/* Loading */}
          {isExtracting && (
            <div style={{
              background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 12, padding: '24px',
              textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10,
            }}>
              <div style={{ fontSize: 28, animation: 'spin 1.5s linear infinite' }}>⚙️</div>
              <div style={{ fontSize: 15, fontWeight: 700, color: '#166534' }}>
                Analyzing Document &amp; Performing OCR Text Extraction...
              </div>
            </div>
          )}

          {extractError && (
            <div style={{ background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 10, padding: '14px 18px', color: '#991b1b', fontSize: 13 }}>
              ❌ {extractError}
            </div>
          )}

          {uploadSuccess && (
            <div style={{ background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 12, padding: '18px', color: '#15803d', fontSize: 14, fontWeight: 700 }}>
              ✓ File filed into SharePoint: <span style={{ fontFamily: 'monospace' }}>{uploadSuccess.path}</span>
            </div>
          )}

          {/* Results Form with Dynamic Fields */}
          {ocrResult && (
            <div style={{ background: '#fff', borderRadius: 14, border: '1px solid #cbd5e1', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', overflow: 'hidden' }}>
              <div style={{
                background: '#f8fafc', borderBottom: '1px solid #e2e8f0', padding: '16px 20px',
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: '#0f172a' }}>{ocrResult.filename}</div>
                  <div style={{ fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginTop: 3 }}>
                    <span>Extracted {ocrResult.text_length} characters • Match: {ocrResult.matched_category_name || 'Standard Taxonomy'}</span>
                    <span>•</span>
                    <span style={{ color: '#0369a1', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                      <span style={{ fontSize: 11 }}>🕒</span>
                      <span>Scanned: {formatScannedTime(extractedAt || new Date().toISOString())}</span>
                    </span>
                  </div>
                </div>
                <span style={{
                  padding: '4px 10px', borderRadius: 20, fontSize: 12, fontWeight: 700,
                  background: ocrResult.confidence >= 0.7 ? '#dcfce7' : '#fef3c7',
                  color: ocrResult.confidence >= 0.7 ? '#15803d' : '#b45309',
                }}>
                  🎯 {(ocrResult.confidence * 100).toFixed(0)}% Match Confidence
                </span>
              </div>

              <div style={{ padding: 24, display: 'flex', flexDirection: 'column', gap: 18 }}>
                {/* Dynamic Category Selector */}
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 6 }}>
                    📂 Document Category (from Templates Module)
                  </label>
                  <select
                    value={selectedMatchCatId || ''}
                    onChange={e => {
                      const cid = Number(e.target.value);
                      setSelectedMatchCatId(cid);
                      const cat = categories.find(c => c.id === cid);
                      if (cat) {
                        setDynamicTags(prev => ({ ...prev, category: cat.name }));
                      }
                    }}
                    style={{
                      width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                      fontSize: 13, fontWeight: 700, color: '#0f172a', background: '#fff',
                    }}
                  >
                    {categories.map(c => (
                      <option key={c.id} value={c.id}>{c.name} ({c.department || 'General'})</option>
                    ))}
                  </select>
                </div>

                {/* Dynamic Tag Fields Form */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
                  {(activeOcrCategory?.tag_fields || [
                    { key: 'vessel', label: 'Vessel Name', type: 'select_vessel', required: true },
                    { key: 'group', label: 'Department', type: 'select_dept', required: true },
                    { key: 'category', label: 'Category', type: 'text', required: true },
                    { key: 'sub_category', label: 'Sub-Category', type: 'text', required: false },
                  ]).map((field: TagFieldDef) => {
                    const val = dynamicTags[field.key] ?? '';

                    if (field.type === 'select_vessel') {
                      return (
                        <div key={field.key}>
                          <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                            🚢 {field.label}
                          </label>
                          <select
                            value={val}
                            onChange={e => setDynamicTags(t => ({ ...t, [field.key]: e.target.value }))}
                            style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff' }}
                          >
                            <option value="">-- Select Vessel --</option>
                            {vesselOptions.map(v => <option key={v} value={v}>{v}</option>)}
                          </select>
                        </div>
                      );
                    }

                    if (field.type === 'select_dept') {
                      return (
                        <div key={field.key}>
                          <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                            🏢 {field.label}
                          </label>
                          <select
                            value={val}
                            onChange={e => setDynamicTags(t => ({ ...t, [field.key]: e.target.value }))}
                            style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff' }}
                          >
                            <option value="Technical & Crewing">Technical & Crewing</option>
                            <option value="Commercial & Chartering">Commercial & Chartering</option>
                            <option value="Insurance">Insurance</option>
                            <option value="Kaizen - Knowledge Bank">Kaizen - Knowledge Bank</option>
                          </select>
                        </div>
                      );
                    }

                    if (field.type === 'select' && field.options && field.options.length > 0) {
                      return (
                        <div key={field.key}>
                          <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                            📑 {field.label}
                          </label>
                          <select
                            value={val}
                            onChange={e => setDynamicTags(t => ({ ...t, [field.key]: e.target.value }))}
                            style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff' }}
                          >
                            <option value="">-- Select {field.label} --</option>
                            {field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        </div>
                      );
                    }

                    return (
                      <div key={field.key}>
                        <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                          🏷️ {field.label}
                        </label>
                        <input
                          type="text"
                          value={val}
                          onChange={e => setDynamicTags(t => ({ ...t, [field.key]: e.target.value }))}
                          style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, boxSizing: 'border-box' }}
                        />
                      </div>
                    );
                  })}
                </div>

                {/* Target SharePoint Path Preview */}
                <div style={{ background: '#f1f5f9', borderRadius: 10, padding: '14px 18px', border: '1px solid #e2e8f0' }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>
                    Resolved Target SharePoint Path
                  </div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: '#0284c7', fontFamily: 'monospace', marginTop: 3 }}>
                    📁 {resolvedOcrPath}
                  </div>
                </div>

                {/* Action Buttons */}
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 12 }}>
                  <button
                    type="button"
                    onClick={() => { setOcrResult(null); setSelectedFile(null); }}
                    style={{ background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
                  >
                    Clear
                  </button>
                  <button
                    type="button"
                    disabled={isUploading}
                    onClick={handleRouteAndUpload}
                    style={{
                      background: isUploading ? '#94a3b8' : '#0078d4', color: '#fff', border: 'none',
                      borderRadius: 8, padding: '10px 24px', fontSize: 13, fontWeight: 700, cursor: isUploading ? 'not-allowed' : 'pointer',
                    }}
                  >
                    {isUploading ? '⏳ Uploading...' : '🚀 Auto-Segregate & Upload to SharePoint'}
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── TAB 3: DOCUMENT CATEGORIES & TAG FIELD MANAGEMENT ── */}
      {activeTab === 'categories' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
            <div>
              <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>
                Document Categories &amp; Metadata Tag Fields
              </h3>
              <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
                The Templates module is the single source of truth for categories, tag schemas, and OCR auto-matching keywords.
              </p>
            </div>
            <button
              onClick={handleOpenNewCategory}
              style={{
                background: '#0078d4', color: '#fff', border: 'none', borderRadius: 8,
                padding: '8px 18px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6,
              }}
            >
              ＋ New Category
            </button>
          </div>

          {/* Categories Table */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
            {loadingCategories ? (
              <div style={{ padding: 30, textAlign: 'center', color: '#64748b' }}>Loading categories...</div>
            ) : categories.length === 0 ? (
              <div style={{ padding: 40, textAlign: 'center', color: '#64748b' }}>No document categories defined yet.</div>
            ) : isMobile ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: 10 }}>
                {categories.map(cat => (
                  <div key={cat.id} style={{ border: '1px solid #e2e8f0', borderRadius: 10, padding: 12 }}>
                    <div style={{ fontWeight: 700, color: '#1e293b', marginBottom: 6 }}>📂 {cat.name}</div>
                    <div style={{ fontSize: 12, marginBottom: 6 }}>{badge('blue', cat.department || 'All Departments')}</div>
                    <div style={{ fontSize: 12, color: '#475569' }}><strong>Tags:</strong> {(cat.tag_fields || []).map(f => f.label).join(', ') || 'None'}</div>
                    <div style={{ fontSize: 12, color: '#475569', marginTop: 6 }}><strong>OCR hints:</strong> {(cat.ocr_hints || []).slice(0, 5).join(', ') || '—'}</div>
                    <details style={{ marginTop: 8 }}>
                      <summary style={{ cursor: 'pointer', color: '#0369a1', fontSize: 12, fontWeight: 700 }}>Path template</summary>
                      <div style={{ fontSize: 11, color: '#0f766e', fontFamily: 'monospace', wordBreak: 'break-all', marginTop: 6 }}>{cat.dms_path_template || '—'}</div>
                    </details>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginTop: 10 }}>
                      <button type="button" onClick={() => handleOpenEditCategory(cat)} style={{ minHeight: 44, background: '#f1f5f9', color: '#334155', border: '1px solid #cbd5e1', borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Edit</button>
                      <button type="button" onClick={() => handleDeleteCategory(cat)} style={{ minHeight: 44, background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca', borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Deactivate</button>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', minWidth: isTablet ? 1040 : 960, borderCollapse: 'collapse', fontSize: 13 }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                    <th style={{ padding: '12px 16px', position: 'sticky', left: 0, zIndex: 2, background: '#f8fafc' }}>Category Name</th>
                    <th style={{ padding: '12px 16px' }}>Department</th>
                    <th style={{ padding: '12px 16px' }}>Dynamic Tag Fields</th>
                    <th style={{ padding: '12px 16px' }}>OCR Keyword Hints</th>
                    <th style={{ padding: '12px 16px' }}>DMS Path Template</th>
                    <th style={{ padding: '12px 16px', textAlign: 'right' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {categories.map(cat => (
                    <tr key={cat.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={{ padding: '14px 16px', fontWeight: 700, color: '#1e293b', position: 'sticky', left: 0, zIndex: 1, background: '#fff' }}>
                        📂 {cat.name}
                      </td>
                      <td style={{ padding: '14px 16px' }}>
                        {badge('blue', cat.department || 'All Departments')}
                      </td>
                      <td style={{ padding: '14px 16px' }}>
                        <span style={{ fontSize: 12, color: '#0369a1', fontWeight: 600 }}>
                          {(cat.tag_fields || []).map(f => f.label).join(', ') || 'None'}
                        </span>
                      </td>
                      <td style={{ padding: '14px 16px', maxWidth: 200 }}>
                        <span style={{ fontSize: 11, color: '#64748b', display: 'inline-block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 190 }}>
                          {(cat.ocr_hints || []).slice(0, 5).join(', ')}...
                        </span>
                      </td>
                      <td style={{ padding: '14px 16px', fontFamily: 'monospace', fontSize: 11, color: '#0f766e' }}>
                        {cat.dms_path_template || '—'}
                      </td>
                      <td style={{ padding: '14px 16px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          onClick={() => handleOpenEditCategory(cat)}
                          style={{
                            background: '#f1f5f9', color: '#334155', border: '1px solid #cbd5e1',
                            borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer', marginRight: 6,
                          }}
                        >
                          ✏️ Edit
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteCategory(cat)}
                          style={{
                            background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca',
                            borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer',
                          }}
                        >
                          🗑️ Deactivate
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── Category Modal (Create / Edit) ── */}
      {isNewCategoryModal && (
        <div
          role="presentation"
          onClick={() => setIsNewCategoryModal(false)}
          style={{
            position: 'fixed', inset: 0, zIndex: 10300,
            background: 'rgba(15,23,42,0.65)', backdropFilter: 'blur(4px)',
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 20,
          }}
        >
          <div
            role="dialog"
            aria-modal="true"
            onClick={e => e.stopPropagation()}
            style={{
              width: isMobile ? '95vw' : 660, maxWidth: '95vw', background: '#fff', borderRadius: 16,
              boxShadow: '0 25px 60px rgba(0,0,0,0.35)', overflow: 'hidden',
              display: 'flex', flexDirection: 'column', maxHeight: '92vh',
            }}
          >
            {/* Header */}
            <div style={{
              padding: '18px 24px', borderBottom: '1px solid #e2e8f0', background: '#f8fafc',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            }}>
              <div style={{ fontSize: 16, fontWeight: 800, color: '#0f172a' }}>
                {editingCategory ? `Edit Category: ${editingCategory.name}` : 'Create New Document Category'}
              </div>
              <button
                type="button"
                onClick={() => setIsNewCategoryModal(false)}
                style={{ background: 'none', border: 'none', fontSize: 20, color: '#94a3b8', cursor: 'pointer' }}
              >
                ✕
              </button>
            </div>

            {/* Body */}
            <div style={{ padding: '20px 24px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 }}>
              {categoryModalError && (
                <div style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#991b1b', padding: '10px 14px', borderRadius: 8, fontSize: 12 }}>
                  ❌ {categoryModalError}
                </div>
              )}

              <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 14 }}>
                {/* Name */}
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                    Category Name *
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. Invoice, Certificate, Contract"
                    value={catName}
                    onChange={e => setCatName(e.target.value)}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, boxSizing: 'border-box' }}
                  />
                </div>

                {/* Department */}
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                    Main Department / Group
                  </label>
                  <select
                    value={catDept}
                    onChange={e => setCatDept(e.target.value)}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff' }}
                  >
                    <option value="Technical & Crewing">Technical & Crewing</option>
                    <option value="Commercial & Chartering">Commercial & Chartering</option>
                    <option value="Insurance">Insurance</option>
                    <option value="Kaizen - Knowledge Bank">Kaizen - Knowledge Bank</option>
                  </select>
                </div>
              </div>

              {/* Path Template */}
              <div>
                <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                  Target DMS Folder Path Template
                </label>
                <input
                  type="text"
                  placeholder="{group}/{vessel}/Drawings and Manuals/{category}/{sub_category}"
                  value={catPathTemplate}
                  onChange={e => setCatPathTemplate(e.target.value)}
                  style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, fontFamily: 'monospace', boxSizing: 'border-box' }}
                />
                <div style={{ fontSize: 11, color: '#64748b', marginTop: 3 }}>
                  Use placeholders matching field keys, e.g. <code style={{ color: '#0284c7' }}>{'{group}'}</code>, <code style={{ color: '#0284c7' }}>{'{vessel}'}</code>, <code style={{ color: '#0284c7' }}>{'{sub_category}'}</code>
                </div>
              </div>

              {/* Tag Fields Builder */}
              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <label style={{ fontSize: 13, fontWeight: 800, color: '#0f172a' }}>
                    🏷️ Associated Tag Fields Schema
                  </label>
                  <button
                    type="button"
                    onClick={() => {
                      setCatTagFields(prev => [
                        ...prev,
                        { key: `field_${prev.length + 1}`, label: `Custom Field ${prev.length + 1}`, type: 'text', required: false, options: null },
                      ]);
                    }}
                    style={{ background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                  >
                    ＋ Add Tag Field
                  </button>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: 180, overflowY: 'auto' }}>
                  {catTagFields.map((f, fIdx) => (
                    <div key={fIdx} style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: isMobile ? 'wrap' : 'nowrap', background: '#f8fafc', padding: 8, borderRadius: 8, border: '1px solid #e2e8f0' }}>
                      <input
                        type="text"
                        placeholder="key (snake_case)"
                        value={f.key}
                        onChange={e => {
                          const val = e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, '_');
                          setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, key: val } : item));
                        }}
                        style={{ width: isMobile ? '100%' : 110, padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }}
                      />
                      <input
                        type="text"
                        placeholder="Label"
                        value={f.label}
                        onChange={e => {
                          const val = e.target.value;
                          setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, label: val } : item));
                        }}
                        style={{ flex: 1, minWidth: isMobile ? '100%' : 120, padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }}
                      />
                      <select
                        value={f.type}
                        onChange={e => {
                          const val = e.target.value as TagFieldDef['type'];
                          setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, type: val } : item));
                        }}
                        style={{ width: isMobile ? '100%' : 120, padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 11, background: '#fff' }}
                      >
                        <option value="text">Text</option>
                        <option value="textarea">Text Area</option>
                        <option value="select_vessel">Vessel Dropdown</option>
                        <option value="select_dept">Dept Dropdown</option>
                        <option value="select">Custom Select</option>
                      </select>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11, color: '#475569', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={f.required}
                          onChange={e => {
                            const checked = e.target.checked;
                            setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, required: checked } : item));
                          }}
                        />
                        Req
                      </label>
                      <button
                        type="button"
                        onClick={() => setCatTagFields(prev => prev.filter((_, i) => i !== fIdx))}
                        style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: 14 }}
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* OCR Keywords */}
              <div>
                <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 4 }}>
                  🔍 OCR Auto-Classification Keyword Hints
                </label>
                <textarea
                  rows={2}
                  placeholder="e.g. invoice, tax invoice, vendor payment, billing, po number, due date"
                  value={catOcrHintsText}
                  onChange={e => setCatOcrHintsText(e.target.value)}
                  style={{ width: '100%', padding: '8px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, boxSizing: 'border-box' }}
                />
                <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>
                  Comma-separated keywords matched against document text during OCR auto-tagging.
                </div>
              </div>
            </div>

            {/* Footer */}
            <div style={{
              padding: '14px 24px', borderTop: '1px solid #e2e8f0', background: '#f8fafc',
              display: 'flex', justifyContent: 'flex-end', gap: 10, flexDirection: isMobile ? 'column' : 'row',
            }}>
              <button
                type="button"
                onClick={() => setIsNewCategoryModal(false)}
                style={{ background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 16px', fontSize: 13, fontWeight: 600, color: '#475569', cursor: 'pointer' }}
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={isSavingCategory}
                onClick={handleSaveCategory}
                style={{
                  background: isSavingCategory ? '#94a3b8' : '#0078d4', color: '#fff', border: 'none',
                  borderRadius: 8, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 22px', fontSize: 13, fontWeight: 700, cursor: isSavingCategory ? 'not-allowed' : 'pointer',
                }}
              >
                {isSavingCategory ? 'Saving...' : 'Save Category'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
