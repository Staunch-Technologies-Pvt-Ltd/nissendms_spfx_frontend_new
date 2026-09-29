/* eslint-disable @typescript-eslint/no-unused-vars */
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge } from '../constants';
import { clay } from '../clayTheme';
import type { DocumentCategory, TagFieldDef } from '../types/ui';
import { OcrStagingQueue } from './OcrStagingQueue';
import { isMobileWidth, isTabletWidth } from '../responsive';

export function renderTemplatesPage(host: VesselEmail): React.ReactElement {
  return <TemplatesPageContainer host={host} />;
}

interface TemplatesPageContainerProps {
  host: VesselEmail;
}

const TemplatesPageContainer: React.FC<TemplatesPageContainerProps> = ({ host }) => {
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const isMobile = isMobileWidth(viewportWidth);
  const isTablet = isTabletWidth(viewportWidth);

  // Active Top Tab: 'queue' | 'categories'
  const [activeTab, setActiveTab] = React.useState<'queue' | 'categories'>('queue');

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
            Manage category taxonomies and review OCR-staged documents.
          </p>
        </div>

        {/* 2-Tab Pill Navigation */}
        <div style={{
          display: 'inline-flex', background: clay.bg, borderRadius: 12, padding: 4,
          border: `1px solid ${clay.accentSoft}`, boxShadow: 'inset 0 1px 2px rgba(0,0,0,0.04)',
          flexWrap: 'wrap',
        }}>
          {/* Tab 1: Queue */}
          <button
            type="button"
            onClick={() => setActiveTab('queue')}
            style={{
              padding: '8px 18px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              background: activeTab === 'queue' ? clay.accentGradient : 'transparent',
              color: activeTab === 'queue' ? '#fff' : clay.textMuted,
              boxShadow: activeTab === 'queue' ? clay.shadowButton : 'none',
              display: 'inline-flex', alignItems: 'center', gap: 8, transition: 'all 0.15s ease',
            }}
          >
            <span>📥</span> Review Staging Queue
            {stagingBadgeCount > 0 && (
              <span style={{
                fontSize: 11, background: activeTab === 'queue' ? '#fff' : clay.accent,
                color: activeTab === 'queue' ? clay.accentDark : '#fff', padding: '1px 7px', borderRadius: 12, fontWeight: 800,
              }}>
                {stagingBadgeCount}
              </span>
            )}
          </button>

          {/* Tab 2: Categories & Templates */}
          <button
            type="button"
            onClick={() => setActiveTab('categories')}
            style={{
              padding: '8px 18px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              background: activeTab === 'categories' ? clay.accentGradient : 'transparent',
              color: activeTab === 'categories' ? '#fff' : clay.textMuted,
              boxShadow: activeTab === 'categories' ? clay.shadowButton : 'none',
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

      {/* ── TAB 2: DOCUMENT CATEGORIES & TAG FIELD MANAGEMENT ── */}
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
                background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: 8,
                padding: '8px 18px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: clay.shadowButton,
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
                  Use placeholders matching field keys, e.g. <code style={{ color: clay.accentDark }}>{'{group}'}</code>, <code style={{ color: clay.accentDark }}>{'{vessel}'}</code>, <code style={{ color: clay.accentDark }}>{'{sub_category}'}</code>
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
                  background: isSavingCategory ? '#94a3b8' : clay.accentGradient, color: '#fff', border: 'none',
                  borderRadius: 8, minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 22px', fontSize: 13, fontWeight: 700, cursor: isSavingCategory ? 'not-allowed' : 'pointer', boxShadow: !isSavingCategory ? clay.shadowButton : 'none',
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
