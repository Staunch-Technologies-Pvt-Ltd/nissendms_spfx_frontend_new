import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { Icon } from '@fluentui/react/lib/Icon';
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
  ApprovalItem, NotificationItem, UserItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import { isMobileWidth } from '../responsive';
import { clay } from '../clayTheme';

export function renderVesselForm(host: VesselEmail, mode: 'create' | 'edit'): React.ReactElement {
  return <VesselFormContent host={host} mode={mode} />;
}

function VesselFormContent({ host, mode }: { host: VesselEmail; mode: 'create' | 'edit' }): React.ReactElement {
  const { form, modalBusy, modalMsg, modalError, formFieldErrors, vessels } = host.state;  const isCreate = mode === 'create';
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));
  const [elapsed, setElapsed] = React.useState(0);
  const [availableSites, setAvailableSites] = React.useState<Array<{ id: string; site_key?: string; name?: string; display_name: string; web_url?: string; is_default?: boolean; is_primary?: boolean }>>([]);
  const [loadingSites, setLoadingSites] = React.useState<boolean>(false);
  const [selectedSiteKey, setSelectedSiteKey] = React.useState(form.site_key || '');
  const [browserPath, setBrowserPath] = React.useState(form.parent_folder_path || '');
  const [browserFolders, setBrowserFolders] = React.useState<Array<{ id: string; name: string; path: string }>>([]);
  const [browserLoading, setBrowserLoading] = React.useState(false);
  const [browserError, setBrowserError] = React.useState('');
  const [parentChosen, setParentChosen] = React.useState(form.parent_folder_path !== undefined);
  const [subfolders, setSubfolders] = React.useState<string[]>(form.subfolders || []);
  const [newSubfolder, setNewSubfolder] = React.useState('');
  const [siteMenuOpen, setSiteMenuOpen] = React.useState(false);
  const siteMenuRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const closeSiteMenu = (event: MouseEvent): void => {
      if (siteMenuRef.current && !siteMenuRef.current.contains(event.target as Node)) {
        setSiteMenuOpen(false);
      }
    };
    document.addEventListener('mousedown', closeSiteMenu);
    return () => document.removeEventListener('mousedown', closeSiteMenu);
  }, []);

  React.useEffect(() => {
    setLoadingSites(true);
    fetch(`${host._base()}/api/sites`, { headers: host._headers() })
      .then(r => r.ok ? r.json() : Promise.reject(r))
      .then(d => {
        const sitesList = d.sites || [];
        setAvailableSites(sitesList);
        if (!form.site_key && sitesList.length > 0) {
          const def = sitesList.find((s: any) => s.is_default || s.is_primary) || sitesList[0];
          const key = def.site_key || def.id;
          setSelectedSiteKey(key);
          host.setState({ form: { ...form, site_key: key, parent_folder_path: undefined, subfolders: form.subfolders || [] } });
        }
      })
      .catch(() => {})
      .finally(() => setLoadingSites(false));
  }, []);

  React.useEffect(() => {
    if (!isCreate || !selectedSiteKey) {
      setBrowserFolders([]);
      return;
    }
    setBrowserLoading(true);
    setBrowserError('');
    fetch(`${host._base()}/api/admin/sites/${encodeURIComponent(selectedSiteKey)}/folders?path=${encodeURIComponent(browserPath)}`, { headers: host._headers() })
      .then(r => r.ok ? r.json() : r.json().then(d => Promise.reject(new Error(d.detail || 'Could not browse SharePoint folders.'))))
      .then(d => setBrowserFolders(d.folders || []))
      .catch((error: Error) => { setBrowserFolders([]); setBrowserError(error.message); })
      .finally(() => setBrowserLoading(false));
  }, [selectedSiteKey, browserPath, isCreate]);

  const chooseSite = (siteKey: string): void => {
    setSelectedSiteKey(siteKey);
    setSiteMenuOpen(false);
    setBrowserPath('');
    setParentChosen(false);
    host.setState({ form: { ...form, site_key: siteKey, parent_folder_path: undefined, subfolders } });
  };

  const chooseParent = (): void => {
    setParentChosen(true);
    host.setState({ form: { ...form, site_key: selectedSiteKey, parent_folder_path: browserPath, subfolders } });
  };

  const addSubfolder = (): void => {
    const value = newSubfolder.trim();
    if (!value || subfolders.some(folder => folder.toLowerCase() === value.toLowerCase())) return;
    const next = [...subfolders, value];
    setSubfolders(next);
    setNewSubfolder('');
    host.setState({ form: { ...form, subfolders: next } });
  };

  React.useEffect(() => {
    if (!modalBusy) {
      setElapsed(0);
      return;
    }
    const timer = setInterval(() => setElapsed(e => e + 1), 1000);
    return () => clearInterval(timer);
  }, [modalBusy]);

  const formatTimer = (sec: number): string => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${m < 10 ? '0' + m : m}:${s < 10 ? '0' + s : s}`;
  };

  const set = (k: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
    host.setState({ form: { ...form, [k]: e.target.value }, modalError: null, formFieldErrors: { ...formFieldErrors, [k]: '' } });

  const setName = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const value = e.target.value;
    const normalized = value.replace(/[ _'\"]+/g, '').toLowerCase();
    const duplicate = normalized.length > 0 && vessels.some(v =>
      v.id !== host.state.selectedVessel?.id &&
      v.name.replace(/[ _'\"]+/g, '').toLowerCase() === normalized
    );
    host.setState({
      form: { ...form, name: value },
      modalError: null,
      formFieldErrors: { ...formFieldErrors, name: duplicate ? 'Vessel name already exists.' : '' },
    });
  };

  const isSuccess = modalMsg && modalMsg.startsWith('🎉');
  const isEditSuccess = !isCreate && !!modalMsg && (modalMsg.startsWith('✅') || modalMsg.startsWith('⏳'));

   const handleClose = () => {
    host.setState({ modal: 'none', modalMsg: null, modalError: null, formFieldErrors: {} });
    if (isSuccess || isEditSuccess) {
      host._loadData();
    }
  };

  const isFormState = !modalBusy && !isSuccess && !isEditSuccess;

  return (
    <div
      style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.55)', backdropFilter: 'blur(3px)', WebkitBackdropFilter: 'blur(3px)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, padding: isMobile ? 10 : 20 }}
      onClick={e => { if (e.target === e.currentTarget && !modalBusy) handleClose(); }}
    >
      <div style={{
        background: clay.surface, borderRadius: 24,
        width: isMobile ? '94vw' : 460, maxWidth: '94vw',
        maxHeight: isMobile ? '92vh' : '86vh',
        display: 'flex', flexDirection: 'column',
        boxShadow: '0 24px 60px -12px rgba(15,23,42,0.35), 0 8px 20px -8px rgba(15,23,42,0.18)',
        border: `1px solid ${clay.accentSoft}`, position: 'relative', overflow: 'hidden',
      }}>

        <div aria-hidden="true" style={{ position: 'absolute', top: 0, left: 0, right: 0, height: 6, background: clay.accentGradient, zIndex: 1 }} />

        {!modalBusy && (
          <button
            onClick={handleClose}
            style={{ position: 'absolute', top: 16, right: 16, width: 28, height: 28, display: 'flex', alignItems: 'center', justifyContent: 'center', background: clay.surfaceRaised, border: `1px solid ${clay.accentSoft}`, borderRadius: 9, fontSize: 14, color: clay.textMuted, cursor: 'pointer', zIndex: 2 }}
            title="Close"
          ><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 14 }} /></button>
        )}

        <div style={{
          padding: isFormState ? (isMobile ? '22px 18px 14px' : '26px 24px 16px') : (isMobile ? '30px 16px 20px' : '30px 28px 22px'),
          overflowY: 'auto', flex: 1, minHeight: 0,
        }}>

        {/* ── STATE 1: PROVISIONING IN PROGRESS (Timer & Progress Bar) ── */}
        {modalBusy && !isCreate ? (
          <div style={{ textAlign: 'center', padding: '28px 0' }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 44 }} /></div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0284c7' }}>
              Updating Vessel
            </h3>
            <p style={{ margin: 0, fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              Saving the updated vessel details. SharePoint folders will not be provisioned.
            </p>
          </div>
        ) : modalBusy ? (
          <div style={{ textAlign: 'center', padding: '12px 0' }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 44 }} /></div>
            <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' }}>
              Creating Vessel…
            </h3>
            <p style={{ margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              Registering <strong>"{form.name || 'Vessel'}"</strong> and creating the selected SharePoint folder path. Please wait.
            </p>

            {/* Live Timer badge */}
            <div style={{
              display: 'inline-flex', alignItems: 'center', gap: 8,
              background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
              borderRadius: 20, padding: '8px 18px', fontSize: 14, fontWeight: 700, marginBottom: 20,
            }}>
              <span style={{ fontSize: 16 }}><Icon iconName="Clock" aria-hidden="true" style={{ fontSize: 16 }} /></span>
              <span>Elapsed Time: {formatTimer(elapsed)}</span>
            </div>

            {/* Animated Progress Bar */}
            <div style={{ background: '#e2e8f0', borderRadius: 10, height: 8, overflow: 'hidden', marginBottom: 20 }}>
              <div style={{
                background: 'linear-gradient(90deg, #0ea5e9, #0284c7, #38bdf8)',
                height: '100%', width: `${Math.min(96, 20 + elapsed * 15)}%`,
                transition: 'width 0.8s ease-out', borderRadius: 10,
              }} />
            </div>

            {/* Step Progress Checklist */}
            <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 }}>
              <div style={{ color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 14 }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 14 }} /></span><span>Vessel record registered in database</span>
              </div>
              <div style={{ color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 14 }}><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 14 }} /></span><span>Creating the vessel folder and selected custom subfolders…</span>
              </div>
              <div style={{ color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 14 }}><Icon iconName="CircleRing" aria-hidden="true" style={{ fontSize: 14 }} /></span><span>Linking category subfolders & permissions</span>
              </div>
            </div>
          </div>
        ) : isSuccess ? (
          /* ── STATE 2: SUCCESS SCREEN ── */
          <div style={{ textAlign: 'center', padding: '12px 0' }}>
            <div style={{ fontSize: 52, marginBottom: 12 }}><Icon iconName="Completed" aria-hidden="true" style={{ fontSize: 52 }} /></div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' }}>
              Vessel Successfully Created!
            </h3>
            <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              <strong>"{form.name}"</strong> has been registered and created at the selected SharePoint location.
            </p>

            <div style={{ background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '14px 16px', marginBottom: 24, fontSize: 12, color: '#065f46', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, marginBottom: 4 }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /> Vessel Ready</div>
              <div>• Registered vessel card added to main grid</div>
              <div>• IMO Number: {form.imo || '—'}</div>
              <div>• Selected SharePoint folder path created</div>
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
        ) : isEditSuccess ? (
          <div style={{ textAlign: 'center', padding: '28px 0' }}>
            <div style={{ fontSize: 52, marginBottom: 12 }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 52 }} /></div>
            <h3 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' }}>
              Vessel Updated Successfully
            </h3>
            <p style={{ margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 }}>
              The vessel details were saved without changing its SharePoint folders.
            </p>
            <button
              onClick={handleClose}
              style={{
                background: 'linear-gradient(135deg, #10b981, #059669)',
                color: '#fff', border: 'none', borderRadius: 10,
                padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
                boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
              }}
            >
              Done
            </button>
          </div>
        ) : (
          /* ── STATE 3: FORM ENTRY ── */
          <>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 18 }}>
              <div style={{ width: 40, height: 40, flexShrink: 0, borderRadius: 12, background: clay.accentGradient, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 19, boxShadow: clay.shadowIcon, border: `1px solid ${clay.accentSoft}` }}>
                <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 19 }} />
              </div>
              <div>
                <div style={{ fontSize: 17, fontWeight: 800, color: clay.text, letterSpacing: '-0.2px', lineHeight: 1.25 }}>{isCreate ? 'Create a new vessel' : 'Update vessel details'}</div>
                <div style={{ fontSize: 12, color: clay.textMuted, marginTop: 2, lineHeight: 1.4 }}>
                  {isCreate ? 'Pick a SharePoint location and optional subfolders.' : 'Update the vessel details below.'}
                </div>
              </div>
            </div>

            {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalMsg}</div>}
            {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalError}</div>}

            <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1.3fr 1fr', gap: 12 }}>
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: clay.text, marginBottom: 5 }}>
                    Vessel name <span style={{ color: '#ef4444' }}>*</span>
                  </label>
                  <input
                    style={{ width: '100%', padding: '9px 12px', borderRadius: 9, border: `1px solid ${formFieldErrors?.name ? '#ef4444' : clay.accentSoft}`, fontSize: 13, outline: 'none', boxSizing: 'border-box', color: clay.text }}
                    value={form.name} onChange={setName} placeholder="e.g. MV Pacific Trader" />
                  {formFieldErrors?.name && (
                    <div style={{ marginTop: 5, fontSize: 11, color: '#dc2626', display: 'flex', alignItems: 'center', gap: 4 }}>
                      <span><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 11 }} /></span><span>{formFieldErrors.name}</span>
                    </div>
                  )}
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: clay.text, marginBottom: 5 }}>
                    IMO number <span style={{ color: '#ef4444' }}>*</span>
                  </label>
                  <input
                    style={{
                      width: '100%', padding: '9px 12px', borderRadius: 9,
                      border: `1px solid ${formFieldErrors?.imo ? '#ef4444' : clay.accentSoft}`,
                      fontSize: 13, outline: 'none', boxSizing: 'border-box', color: clay.text,
                    }}
                    value={form.imo}
                    onChange={e => {
                      const val = e.target.value.replace(/\D/g, '').slice(0, 7);
                      const duplicate = val.length > 0 && vessels.some(v =>
                        v.id !== host.state.selectedVessel?.id && (v.imo || '').trim() === val
                      );
                      host.setState({
                        form: { ...form, imo: val },
                        modalError: null,
                        formFieldErrors: { ...formFieldErrors, imo: duplicate ? 'A vessel with that IMO number already exists.' : '' },
                      });
                    }}
                    placeholder="7 digits, e.g. 9074729"
                    maxLength={7}
                    inputMode="numeric" />
                  {formFieldErrors?.imo && (
                    <div style={{ marginTop: 5, fontSize: 11, color: '#dc2626', display: 'flex', alignItems: 'center', gap: 4 }}>
                      <span><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 11 }} /></span><span>{formFieldErrors.imo}</span>
                    </div>
                  )}
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 12 }}>
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: clay.text, marginBottom: 5 }}>Ship yard name</label>
                  <input
                    style={{ width: '100%', padding: '9px 12px', borderRadius: 9, border: `1px solid ${clay.accentSoft}`, fontSize: 13, outline: 'none', boxSizing: 'border-box', color: clay.text }}
                    value={form.shipyard} onChange={set('shipyard')} placeholder="e.g. Hyundai Heavy Industries" />
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: clay.text, marginBottom: 5 }}>Hull number</label>
                  <input
                    style={{ width: '100%', padding: '9px 12px', borderRadius: 9, border: `1px solid ${clay.accentSoft}`, fontSize: 13, outline: 'none', boxSizing: 'border-box', color: clay.text }}
                    value={form.hull_number} onChange={set('hull_number')} placeholder="e.g. H2456" />
                </div>
              </div>

              <div>
                <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: clay.text, marginBottom: 5 }}>Vessel type</label>
                <select
                  style={{ width: '100%', padding: '9px 12px', borderRadius: 9, border: `1px solid ${clay.accentSoft}`, fontSize: 13, outline: 'none', boxSizing: 'border-box', background: clay.surface, color: form.vessel_type ? clay.text : clay.textMuted, appearance: 'auto' }}
                  value={form.vessel_type} onChange={set('vessel_type')}>
                  <option value="">Select a type...</option>
                  {VESSEL_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>

              {isCreate && (
                <div style={{ background: clay.surfaceRaised, borderRadius: 12, border: `1px solid ${clay.accentSoft}`, padding: 12, boxSizing: 'border-box', display: 'flex', flexDirection: 'column', gap: 12 }}>
                  <div>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 7, fontSize: 12, fontWeight: 700, color: clay.text, marginBottom: 6 }}>
                      <span style={{ width: 17, height: 17, borderRadius: '50%', background: clay.accentGradient, color: '#fff', fontSize: 10, fontWeight: 800, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>1</span>
                      SharePoint site
                    </label>
                    <div ref={siteMenuRef} style={{ position: 'relative' }}>
                      <button
                        type="button"
                        disabled={modalBusy || loadingSites}
                        aria-haspopup="listbox"
                        aria-expanded={siteMenuOpen}
                        onClick={() => setSiteMenuOpen(open => !open)}
                        style={{ width: '100%', minHeight: 42, padding: '10px 38px 10px 12px', boxSizing: 'border-box', border: `1px solid ${siteMenuOpen ? clay.accent : clay.accentSoft}`, borderRadius: 8, background: clay.surface, color: selectedSiteKey ? clay.text : clay.textMuted, textAlign: 'left', cursor: modalBusy || loadingSites ? 'not-allowed' : 'pointer', position: 'relative', fontSize: 13 }}
                      >
                        {availableSites.find(site => (site.site_key || site.id) === selectedSiteKey)?.display_name || selectedSiteKey || (loadingSites ? 'Loading sites...' : 'Select a SharePoint site')}
                        <span aria-hidden="true" style={{ position: 'absolute', right: 13, top: '50%', transform: `translateY(-50%) rotate(${siteMenuOpen ? 180 : 0}deg)`, color: clay.accentDark, transition: 'transform 0.15s ease' }}><Icon iconName="ChevronDown" aria-hidden="true" style={{ fontSize: 12 }} /></span>
                      </button>
                      {siteMenuOpen && !modalBusy && !loadingSites && (
                        <div role="listbox" aria-label="SharePoint sites" style={{ position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, zIndex: 20, maxHeight: 190, overflowY: 'auto', background: clay.surface, border: `1px solid ${clay.accentSoft}`, borderRadius: 8, padding: 4 }}>
                          {availableSites.map(site => {
                            const key = site.site_key || site.id;
                            const label = site.display_name || site.name || key;
                            const selected = key === selectedSiteKey;
                            return (
                              <button
                                type="button"
                                role="option"
                                aria-selected={selected}
                                key={key}
                                onClick={() => chooseSite(key)}
                                style={{ width: '100%', minHeight: 36, border: 0, borderRadius: 6, padding: '8px 10px', boxSizing: 'border-box', background: selected ? clay.accentSoft : 'transparent', color: selected ? clay.accentDeep : clay.text, textAlign: 'left', cursor: 'pointer', fontSize: 13, fontWeight: selected ? 700 : 500 }}
                              >
                                {selected && <span style={{ marginRight: 6, color: clay.accentDark }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /></span>}{label}
                              </button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>

                  {selectedSiteKey && (
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
                        <label style={{ display: 'flex', alignItems: 'center', gap: 7, fontSize: 12, fontWeight: 700, color: clay.text }}>
                          <span style={{ width: 17, height: 17, borderRadius: '50%', background: clay.accentGradient, color: '#fff', fontSize: 10, fontWeight: 800, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>2</span>
                          Choose parent folder
                        </label>
                        <button type="button" onClick={() => { setBrowserPath(''); }} disabled={!browserPath || browserLoading} style={{ border: 0, background: 'transparent', color: '#0369a1', cursor: 'pointer', fontSize: 12 }}>Root</button>
                      </div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 4, fontSize: 12, color: '#475569', marginBottom: 8 }}>
                        <button type="button" onClick={() => setBrowserPath('')} style={{ border: 0, background: 'transparent', padding: 0, color: '#0369a1', cursor: 'pointer', fontSize: 12, fontWeight: 600 }}>Root</button>
                        {browserPath.split('/').filter(Boolean).map((segment, index, segments) => (
                          <React.Fragment key={`${segment}-${index}`}>
                            <span>/</span>
                            <button type="button" onClick={() => setBrowserPath(segments.slice(0, index + 1).join('/'))} style={{ border: 0, background: 'transparent', padding: 0, color: index === segments.length - 1 ? '#0f172a' : '#0369a1', cursor: 'pointer', fontSize: 12, fontWeight: index === segments.length - 1 ? 700 : 600 }}>{segment}</button>
                          </React.Fragment>
                        ))}
                      </div>
                      {browserError && <div style={{ color: '#b91c1c', fontSize: 12, marginBottom: 8 }}>{browserError}</div>}
                      {browserLoading ? <div style={{ padding: 12, color: clay.textMuted, fontSize: 12 }}>Loading folders...</div> : (
                        <div style={{ border: `1px solid ${clay.accentSoft}`, borderRadius: 8, background: clay.surface, maxHeight: 180, overflowY: 'auto' }}>
                          {browserFolders.length === 0 ? <div style={{ padding: 12, color: clay.textMuted, fontSize: 12 }}>No subfolders at this level.</div> : browserFolders.map(folder => (
                            <button type="button" key={folder.id} onClick={() => setBrowserPath(folder.path)} style={{ width: '100%', textAlign: 'left', padding: '9px 12px', boxSizing: 'border-box', border: 0, borderBottom: `1px solid ${clay.accentSoft}`, background: clay.surface, color: clay.text, cursor: 'pointer' }}><Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 12 }} /> {folder.name}</button>
                          ))}
                        </div>
                      )}
                      <button type="button" onClick={chooseParent} disabled={modalBusy} style={{ marginTop: 8, padding: '8px 12px', border: `1px solid ${form.parent_folder_path === browserPath ? '#059669' : '#0284c7'}`, borderRadius: 7, background: form.parent_folder_path === browserPath ? '#dcfce7' : '#eff6ff', color: form.parent_folder_path === browserPath ? '#15803d' : '#0369a1', cursor: 'pointer', fontWeight: 600, fontSize: 12 }}>
                        {form.parent_folder_path === browserPath ? <><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /> Parent folder selected</> : 'Use this folder as parent'}
                      </button>
                    </div>
                  )}

                  {selectedSiteKey && (
                    <div>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 7, fontSize: 12, fontWeight: 700, color: clay.text, marginBottom: 6 }}>
                        <span style={{ width: 17, height: 17, borderRadius: '50%', background: clay.accentGradient, color: '#fff', fontSize: 10, fontWeight: 800, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>3</span>
                        Folders inside {form.name || 'the vessel'}
                      </label>
                      <div style={{ fontSize: 11, color: clay.textMuted, marginBottom: 8 }}>
                        Add flat folders or nested paths such as Certificates/Statutory/Reports. They will be created inside the new vessel folder.
                        {!parentChosen && ' Select the parent folder above before creating the vessel.'}
                      </div>
                      <div style={{ display: 'flex', gap: 8 }}>
                        <input value={newSubfolder} onChange={e => setNewSubfolder(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); addSubfolder(); } }} placeholder="e.g. Certificates/Statutory/Reports" style={{ flex: 1, minWidth: 0, padding: '9px 10px', boxSizing: 'border-box', border: `1px solid ${clay.accentSoft}`, borderRadius: 7 }} />
                        <button type="button" onClick={addSubfolder} disabled={modalBusy} style={{ padding: '8px 12px', border: '1px solid #0284c7', borderRadius: 7, background: '#eff6ff', color: '#0369a1', cursor: modalBusy ? 'not-allowed' : 'pointer', fontWeight: 600 }}>+ Add folder</button>
                      </div>
                      {subfolders.length > 0 && <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 9 }}>{subfolders.map(folder => <span key={folder} style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: '#e0f2fe', color: '#075985', borderRadius: 14, padding: '4px 9px', fontSize: 12 }}>{folder}<button type="button" onClick={() => { const next = subfolders.filter(value => value !== folder); setSubfolders(next); host.setState({ form: { ...form, subfolders: next } }); }} style={{ border: 0, background: 'transparent', color: '#075985', cursor: 'pointer', padding: 0 }}>×</button></span>)}</div>}
                    </div>
                  )}
                </div>
              )}
            </div>
          </>
        )}
        </div>

        {isFormState && (
          <div style={{
            display: 'flex', justifyContent: 'flex-end', flexDirection: isMobile ? 'column' : 'row', gap: 10,
            padding: isMobile ? '14px 18px' : '14px 24px', borderTop: `1px solid ${clay.accentSoft}`,
            background: clay.surfaceRaised, flexShrink: 0,
          }}>
            <button
              style={{ background: 'transparent', border: `1px solid ${clay.accentSoft}`, borderRadius: 8, minHeight: 40, width: isMobile ? '100%' : 'auto', boxSizing: 'border-box', padding: '9px 18px', fontSize: 13, fontWeight: 500, color: clay.textMuted, cursor: 'pointer' }}
              onClick={handleClose}>Cancel</button>
            <button
              style={{ background: clay.accentGradient, color: '#fff', border: 'none', borderRadius: clay.radiusButton, minHeight: 40, width: isMobile ? '100%' : 'auto', boxSizing: 'border-box', padding: '9px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer', boxShadow: clay.shadowButton }}
              onClick={isCreate ? host._submitCreate : host._submitEdit}>
              {isCreate ? 'Create Vessel' : 'Update Vessel'}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
