/* eslint-disable @typescript-eslint/no-unused-vars */
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge } from '../constants';
import { clay } from '../clayTheme';
import type { DocumentCategory, TagFieldDef } from '../types/ui';
import { OcrStagingQueue } from './OcrStagingQueue';
import { isMobileWidth, isTabletWidth } from '../responsive';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  DmsPageHeader, dmsBtn, dmsRowBtn, dmsControlStyle, DMS_ON_ACCENT,
  DMS_TABLE_CARD, DMS_TABLE, DMS_TH, DMS_TR, DMS_TD, DMS_TD_NAME,
} from '../dmsDesignSystem';

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
      <DmsPageHeader title="Templates, Categories & OCR Staging" subtitle="Manage category taxonomies and review OCR-staged documents.">
        {/* 2-Tab Pill Navigation */}
        <div style={{
          display: 'inline-flex', background: 'var(--vdms-surface-alt)', borderRadius: 12, padding: 4,
          border: '1px solid var(--vdms-line)', boxShadow: 'inset 0 1px 2px rgba(0,0,0,0.04)',
          flexWrap: 'wrap',
        }}>
          {/* Tab 1: Queue */}
          <button
            type="button"
            onClick={() => setActiveTab('queue')}
            style={{
              padding: '8px 18px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 700, cursor: 'pointer',
              background: activeTab === 'queue' ? clay.accentGradient : 'transparent',
              color: activeTab === 'queue' ? DMS_ON_ACCENT : 'var(--vdms-text-muted)',
              boxShadow: activeTab === 'queue' ? clay.shadowButton : 'none',
              display: 'inline-flex', alignItems: 'center', gap: 8, transition: 'all 0.15s ease',
            }}
          >
            <Icon iconName="Inbox" aria-hidden="true" style={{ fontSize: 13 }} /> Review Staging Queue
            {stagingBadgeCount > 0 && (
              <span style={{
                fontSize: 11, background: activeTab === 'queue' ? DMS_ON_ACCENT : clay.accent,
                color: activeTab === 'queue' ? clay.accentDark : DMS_ON_ACCENT, padding: '1px 7px', borderRadius: 12, fontWeight: 800,
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
              color: activeTab === 'categories' ? DMS_ON_ACCENT : 'var(--vdms-text-muted)',
              boxShadow: activeTab === 'categories' ? clay.shadowButton : 'none',
              display: 'inline-flex', alignItems: 'center', gap: 8, transition: 'all 0.15s ease',
            }}
          >
            <Icon iconName="Settings" aria-hidden="true" style={{ fontSize: 13 }} /> Document Categories &amp; Tags
          </button>
        </div>
      </DmsPageHeader>

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
              <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: 'var(--vdms-text)' }}>
                Document Categories &amp; Metadata Tag Fields
              </h3>
              <p style={{ margin: '4px 0 0', fontSize: 13, color: 'var(--vdms-text-muted)' }}>
                The Templates module is the single source of truth for categories, tag schemas, and OCR auto-matching keywords.
              </p>
            </div>
            <button
              onClick={handleOpenNewCategory}
              style={dmsBtn('primary', true)}
            >
              <Icon iconName="Add" aria-hidden="true" style={{ fontSize: 13 }} /> New Category
            </button>
          </div>

          {/* Categories Table */}
          <div style={DMS_TABLE_CARD}>
            {loadingCategories ? (
              <div style={{ padding: 30, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>Loading categories...</div>
            ) : categories.length === 0 ? (
              <div style={{ padding: 40, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>No document categories defined yet.</div>
            ) : isMobile ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: 10 }}>
                {categories.map(cat => (
                  <div key={cat.id} style={{ border: '1px solid var(--vdms-line)', borderRadius: 10, padding: 12 }}>
                    <div style={{ fontWeight: 700, color: 'var(--vdms-text)', marginBottom: 6 }}><Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 13 }} /> {cat.name}</div>
                    <div style={{ fontSize: 12, marginBottom: 6 }}>{badge('blue', cat.department || 'All Departments')}</div>
                    <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}><strong>Tags:</strong> {(cat.tag_fields || []).map(f => f.label).join(', ') || 'None'}</div>
                    <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 6 }}><strong>OCR hints:</strong> {(cat.ocr_hints || []).slice(0, 5).join(', ') || '—'}</div>
                    <details style={{ marginTop: 8 }}>
                      <summary style={{ cursor: 'pointer', color: clay.accent, fontSize: 12, fontWeight: 700 }}>Path template</summary>
                      <div style={{ fontSize: 11, color: clay.accentDark, fontFamily: 'monospace', wordBreak: 'break-all', marginTop: 6 }}>{cat.dms_path_template || '—'}</div>
                    </details>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginTop: 10 }}>
                      <button type="button" onClick={() => handleOpenEditCategory(cat)} style={{ ...dmsRowBtn('plain'), minHeight: 44 }}>Edit</button>
                      <button type="button" onClick={() => handleDeleteCategory(cat)} style={{ ...dmsRowBtn('danger'), minHeight: 44 }}>Deactivate</button>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ overflowX: 'auto' }}>
              <table style={{ ...DMS_TABLE, minWidth: isTablet ? 1040 : 960 }}>
                <thead>
                  <tr>
                    <th style={{ ...DMS_TH, position: 'sticky', left: 0, zIndex: 2, background: 'var(--vdms-surface-alt)' }}>Category Name</th>
                    <th style={DMS_TH}>Department</th>
                    <th style={DMS_TH}>Dynamic Tag Fields</th>
                    <th style={DMS_TH}>OCR Keyword Hints</th>
                    <th style={DMS_TH}>DMS Path Template</th>
                    <th style={{ ...DMS_TH, textAlign: 'right' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {categories.map(cat => (
                    <tr key={cat.id} style={DMS_TR}>
                      <td style={{ ...DMS_TD_NAME, position: 'sticky', left: 0, zIndex: 1, background: 'var(--vdms-surface)' }}>
                        <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 13 }} /> {cat.name}
                      </td>
                      <td style={DMS_TD}>
                        {badge('blue', cat.department || 'All Departments')}
                      </td>
                      <td style={DMS_TD}>
                        <span style={{ fontSize: 12, color: clay.accent, fontWeight: 600 }}>
                          {(cat.tag_fields || []).map(f => f.label).join(', ') || 'None'}
                        </span>
                      </td>
                      <td style={{ ...DMS_TD, maxWidth: 200 }}>
                        <span style={{ fontSize: 11, color: 'var(--vdms-text-muted)', display: 'inline-block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 190 }}>
                          {(cat.ocr_hints || []).slice(0, 5).join(', ')}...
                        </span>
                      </td>
                      <td style={{ ...DMS_TD, fontFamily: 'monospace', fontSize: 11, color: clay.accentDark }}>
                        {cat.dms_path_template || '—'}
                      </td>
                      <td style={{ ...DMS_TD, textAlign: 'right', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          onClick={() => handleOpenEditCategory(cat)}
                          style={{ ...dmsRowBtn('plain'), marginRight: 6 }}
                        >
                          <Icon iconName="Edit" aria-hidden="true" style={{ fontSize: 11 }} /> Edit
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteCategory(cat)}
                          style={dmsRowBtn('danger')}
                        >
                          <Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 11 }} /> Deactivate
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
              width: isMobile ? '95vw' : 660, maxWidth: '95vw', background: 'var(--vdms-surface)', borderRadius: 16,
              boxShadow: clay.shadowRaised, overflow: 'hidden', border: '1px solid var(--vdms-line)',
              display: 'flex', flexDirection: 'column', maxHeight: '92vh',
            }}
          >
            {/* Header */}
            <div style={{
              padding: '18px 24px', borderBottom: '1px solid var(--vdms-line)', background: 'var(--vdms-surface-alt)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            }}>
              <div style={{ fontSize: 16, fontWeight: 800, color: 'var(--vdms-text)' }}>
                {editingCategory ? `Edit Category: ${editingCategory.name}` : 'Create New Document Category'}
              </div>
              <button
                type="button"
                onClick={() => setIsNewCategoryModal(false)}
                style={{ background: 'none', border: 'none', fontSize: 20, color: 'var(--vdms-text-faint)', cursor: 'pointer' }}
              >
                <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 20 }} />
              </button>
            </div>

            {/* Body */}
            <div style={{ padding: '20px 24px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 }}>
              {categoryModalError && (
                <div style={{ background: clay.pillDangerBg, border: '1px solid var(--vdms-line)', color: clay.pillDangerText, padding: '10px 14px', borderRadius: 8, fontSize: 12 }}>
                  <Icon iconName="ErrorBadge" aria-hidden="true" style={{ fontSize: 12 }} /> {categoryModalError}
                </div>
              )}

              <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 14 }}>
                {/* Name */}
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: 'var(--vdms-text-secondary)', marginBottom: 4 }}>
                    Category Name *
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. Invoice, Certificate, Contract"
                    value={catName}
                    onChange={e => setCatName(e.target.value)}
                    style={{ ...dmsControlStyle(), width: '100%', fontSize: 13, boxSizing: 'border-box' }}
                  />
                </div>

                {/* Department */}
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: 'var(--vdms-text-secondary)', marginBottom: 4 }}>
                    Main Department / Group
                  </label>
                  <select
                    value={catDept}
                    onChange={e => setCatDept(e.target.value)}
                    style={{ ...dmsControlStyle(), width: '100%', fontSize: 13 }}
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
                <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: 'var(--vdms-text-secondary)', marginBottom: 4 }}>
                  Target DMS Folder Path Template
                </label>
                <input
                  type="text"
                  placeholder="{group}/{vessel}/Drawings and Manuals/{category}/{sub_category}"
                  value={catPathTemplate}
                  onChange={e => setCatPathTemplate(e.target.value)}
                  style={{ ...dmsControlStyle(), width: '100%', fontSize: 13, fontFamily: 'monospace', boxSizing: 'border-box' }}
                />
                <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)', marginTop: 3 }}>
                  Use placeholders matching field keys, e.g. <code style={{ color: clay.accentDark }}>{'{group}'}</code>, <code style={{ color: clay.accentDark }}>{'{vessel}'}</code>, <code style={{ color: clay.accentDark }}>{'{sub_category}'}</code>
                </div>
              </div>

              {/* Tag Fields Builder */}
              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <label style={{ fontSize: 13, fontWeight: 800, color: 'var(--vdms-text)' }}>
                    <Icon iconName="Tag" aria-hidden="true" style={{ fontSize: 13 }} /> Associated Tag Fields Schema
                  </label>
                  <button
                    type="button"
                    onClick={() => {
                      setCatTagFields(prev => [
                        ...prev,
                        { key: `field_${prev.length + 1}`, label: `Custom Field ${prev.length + 1}`, type: 'text', required: false, options: null },
                      ]);
                    }}
                    style={dmsRowBtn('accent')}
                  >
                    <Icon iconName="Add" aria-hidden="true" style={{ fontSize: 11 }} /> Add Tag Field
                  </button>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: 180, overflowY: 'auto' }}>
                  {catTagFields.map((f, fIdx) => (
                    <div key={fIdx} style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: isMobile ? 'wrap' : 'nowrap', background: 'var(--vdms-surface-alt)', padding: 8, borderRadius: 8, border: '1px solid var(--vdms-line)' }}>
                      <input
                        type="text"
                        placeholder="key (snake_case)"
                        value={f.key}
                        onChange={e => {
                          const val = e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, '_');
                          setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, key: val } : item));
                        }}
                        style={{ ...dmsControlStyle(), width: isMobile ? '100%' : 110, padding: '5px 8px', borderRadius: 6 }}
                      />
                      <input
                        type="text"
                        placeholder="Label"
                        value={f.label}
                        onChange={e => {
                          const val = e.target.value;
                          setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, label: val } : item));
                        }}
                        style={{ ...dmsControlStyle(), flex: 1, minWidth: isMobile ? '100%' : 120, padding: '5px 8px', borderRadius: 6 }}
                      />
                      <select
                        value={f.type}
                        onChange={e => {
                          const val = e.target.value as TagFieldDef['type'];
                          setCatTagFields(prev => prev.map((item, i) => i === fIdx ? { ...item, type: val } : item));
                        }}
                        style={{ ...dmsControlStyle(), width: isMobile ? '100%' : 120, padding: '5px 8px', borderRadius: 6, fontSize: 11 }}
                      >
                        <option value="text">Text</option>
                        <option value="textarea">Text Area</option>
                        <option value="select_vessel">Vessel Dropdown</option>
                        <option value="select_dept">Dept Dropdown</option>
                        <option value="select">Custom Select</option>
                      </select>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11, color: 'var(--vdms-text-muted)', cursor: 'pointer' }}>
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
                        style={{ background: 'none', border: 'none', color: clay.pillDangerText, cursor: 'pointer', fontSize: 14 }}
                      >
                        <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 14 }} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* OCR Keywords */}
              <div>
                <label style={{ display: 'block', fontSize: 12, fontWeight: 700, color: 'var(--vdms-text-secondary)', marginBottom: 4 }}>
                  <Icon iconName="Search" aria-hidden="true" style={{ fontSize: 12 }} /> OCR Auto-Classification Keyword Hints
                </label>
                <textarea
                  rows={2}
                  placeholder="e.g. invoice, tax invoice, vendor payment, billing, po number, due date"
                  value={catOcrHintsText}
                  onChange={e => setCatOcrHintsText(e.target.value)}
                  style={{ ...dmsControlStyle(), width: '100%', boxSizing: 'border-box' }}
                />
                <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                  Comma-separated keywords matched against document text during OCR auto-tagging.
                </div>
              </div>
            </div>

            {/* Footer */}
            <div style={{
              padding: '14px 24px', borderTop: '1px solid var(--vdms-line)', background: 'var(--vdms-surface-alt)',
              display: 'flex', justifyContent: 'flex-end', gap: 10, flexDirection: isMobile ? 'column' : 'row',
            }}>
              <button
                type="button"
                onClick={() => setIsNewCategoryModal(false)}
                style={{ ...dmsBtn('secondary', true), minHeight: 44, width: isMobile ? '100%' : 'auto' }}
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={isSavingCategory}
                onClick={handleSaveCategory}
                style={{ ...dmsBtn('primary', !isSavingCategory), minHeight: 44, width: isMobile ? '100%' : 'auto' }}
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
