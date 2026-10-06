import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { VesselSuggestion } from '../types/ui';
import { isMobileWidth } from '../responsive';
import { Icon } from '@fluentui/react/lib/Icon';

// ── Vessel Suggestions Modal ───────────────────────────────────────────────
// Full-screen wizard that:
//   Step 1 (extract)  — animates through extracting info from uploaded file names
//   Step 2 (match)    — shows Term Store match result with confidence badge
//   Step 3 (review)   — editable form: IMO, Shipyard, Hull No., Vessel Type, Vessel Name
//   Step 4 (done)     — success screen

const VESSEL_TYPES = [
  '', 'Bulk Carrier', 'VLCC', 'Suezmax Tanker', 'Aframax Tanker', 'Product Tanker',
  'LNG Carrier', 'LPG Carrier', 'Chemical Tanker', 'Tanker', 'Container Ship',
  'RoRo / Car Carrier', 'General Cargo', 'Offshore Vessel', 'Passenger / Ferry', 'Other',
];

const OVERLAY_STYLE: React.CSSProperties = {
  position: 'fixed', inset: 0, zIndex: 110000,
  background: 'rgba(8,14,44,0.72)',
  backdropFilter: 'blur(12px)',
  display: 'flex', alignItems: 'center', justifyContent: 'center',
  padding: 20,
  fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif",
};

const CARD_STYLE: React.CSSProperties = {
  background: 'linear-gradient(145deg, #0f172a 0%, #1e293b 100%)',
  border: '1px solid rgba(99,179,237,0.25)',
  borderRadius: 24,
  boxShadow: '0 32px 96px rgba(0,0,0,0.6), 0 0 0 1px rgba(255,255,255,0.04)',
  width: '100%',
  maxWidth: 600,
  position: 'relative',
  overflow: 'hidden',
};

const GLASS_PANEL: React.CSSProperties = {
  background: 'rgba(255,255,255,0.04)',
  border: '1px solid rgba(255,255,255,0.08)',
  borderRadius: 12,
  padding: '14px 16px',
};

function confidenceColor(c: number): string {
  if (c >= 0.75) return '#4ade80';
  if (c >= 0.40) return '#facc15';
  return '#f87171';
}

function confidenceLabel(c: number): string {
  if (c >= 0.75) return 'High confidence';
  if (c >= 0.40) return 'Medium confidence';
  return 'Low confidence';
}

function StepIndicator({ step }: { step: string }): React.ReactElement {
  const steps = [
    { id: 'extract', label: 'Extract' },
    { id: 'match', label: 'Match' },
    { id: 'review', label: 'Review' },
    { id: 'done', label: 'Done' },
  ];
  const activeIdx = steps.findIndex(s => s.id === step);
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 0, marginBottom: 28 }}>
      {steps.map((s, i) => {
        const isActive = s.id === step;
        const isDone = i < activeIdx;
        return (
          <React.Fragment key={s.id}>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6 }}>
              <div style={{
                width: 32, height: 32, borderRadius: '50%',
                background: isDone ? '#4ade80' : isActive ? 'linear-gradient(135deg, #3b82f6, #8b5cf6)' : 'rgba(255,255,255,0.08)',
                border: isActive ? '2px solid rgba(99,179,237,0.6)' : '2px solid transparent',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: 13, fontWeight: 700, color: '#fff',
                boxShadow: isActive ? '0 0 16px rgba(59,130,246,0.5)' : 'none',
                transition: 'all 0.3s ease',
              }}>
                {isDone ? <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> : i + 1}
              </div>
              <span style={{ fontSize: 10, color: isActive ? '#93c5fd' : isDone ? '#86efac' : 'rgba(255,255,255,0.4)', fontWeight: isActive ? 600 : 400 }}>
                {s.label}
              </span>
            </div>
            {i < steps.length - 1 && (
              <div style={{
                flex: 1, height: 2, margin: '0 4px', marginBottom: 20,
                background: isDone ? 'linear-gradient(90deg,#4ade80,#22c55e)' : 'rgba(255,255,255,0.1)',
                transition: 'background 0.4s ease',
              }} />
            )}
          </React.Fragment>
        );
      })}
    </div>
  );
}

function ExtractRow({ label, value, delay }: { label: string; value: string; delay: number }): React.ReactElement {
  const [visible, setVisible] = React.useState(false);
  React.useEffect(() => {
    const t = setTimeout(() => setVisible(true), delay);
    return () => clearTimeout(t);
  }, [delay]);
  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: 12, padding: '8px 0',
      opacity: visible ? 1 : 0, transform: visible ? 'translateX(0)' : 'translateX(-12px)',
      transition: 'all 0.35s ease', borderBottom: '1px solid rgba(255,255,255,0.05)',
    }}>
      <span style={{ fontSize: 11, color: 'rgba(148,163,184,1)', width: 120, flexShrink: 0 }}>{label}</span>
      <span style={{
        fontSize: 13, fontWeight: 600, color: value ? '#e2e8f0' : 'rgba(100,116,139,1)',
        fontFamily: value ? "'JetBrains Mono', monospace" : 'inherit',
      }}>
        {value || 'Not detected'}
      </span>
      {value && <span style={{ marginLeft: 'auto', fontSize: 11, color: '#4ade80' }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 11 }} /></span>}
    </div>
  );
}

function Field({
  label, value, onChange, type = 'text', options, required, hint,
}: {
  label: string; value: string; onChange: (v: string) => void;
  type?: 'text' | 'select'; options?: string[]; required?: boolean; hint?: string;
}): React.ReactElement {
  const inputStyle: React.CSSProperties = {
    width: '100%', boxSizing: 'border-box',
    background: 'rgba(255,255,255,0.06)', border: '1.5px solid rgba(99,179,237,0.2)',
    borderRadius: 10, padding: '10px 14px', color: '#e2e8f0', fontSize: 14,
    outline: 'none', transition: 'border-color 0.2s',
    fontFamily: "'Segoe UI', sans-serif",
  };
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: 'block', fontSize: 11, fontWeight: 600, color: '#94a3b8', marginBottom: 5, letterSpacing: '0.05em', textTransform: 'uppercase' }}>
        {label}{required && <span style={{ color: '#f87171', marginLeft: 3 }}>*</span>}
      </label>
      {type === 'select' ? (
        <select
          value={value}
          onChange={e => onChange(e.target.value)}
          style={{ ...inputStyle, appearance: 'none', cursor: 'pointer' }}
        >
          {(options || []).map(o => <option key={o} value={o} style={{ background: '#1e293b' }}>{o || 'Select type…'}</option>)}
        </select>
      ) : (
        <input
          type="text"
          value={value}
          onChange={e => onChange(e.target.value)}
          style={inputStyle}
          onFocus={e => { e.currentTarget.style.borderColor = 'rgba(99,179,237,0.6)'; }}
          onBlur={e => { e.currentTarget.style.borderColor = 'rgba(99,179,237,0.2)'; }}
        />
      )}
      {hint && <div style={{ fontSize: 10, color: '#64748b', marginTop: 4 }}>{hint}</div>}
    </div>
  );
}

export function renderVesselSuggestionsModal(host: VesselEmail): React.ReactElement | null {
  const dlg = host.state.vesselSuggestionDialog;
  if (!dlg || !dlg.open) return null;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  const {
    step, extracting, suggestion, editedSuggestion, creating, error, files,
    allSuggestions = [], selectedSuggestionIndex = 0,
  } = dlg;

  const close = (): void => { host.setState({ vesselSuggestionDialog: null }); };

  const setEdited = (patch: Partial<VesselSuggestion>): void => {
    host.setState(prev => ({
      vesselSuggestionDialog: prev.vesselSuggestionDialog
        ? { ...prev.vesselSuggestionDialog, editedSuggestion: { ...prev.vesselSuggestionDialog.editedSuggestion!, ...patch } }
        : null,
    }));
  };

  const useExisting = (): void => {
    if (!suggestion?.matchedExisting) return;
    void host._applyVesselSuggestionUseExisting();
  };

  const hasMatch = !!suggestion?.matchedExisting;
  const confPct = Math.round((suggestion?.confidence || 0) * 100);
  const formatLastScanned = (ts?: number): string => {
    if (!ts) return 'Unknown';
    try {
      return new Date(ts).toLocaleString();
    } catch {
      return 'Unknown';
    }
  };

  return (
    <div style={OVERLAY_STYLE} onClick={e => { if (e.target === e.currentTarget) close(); }}>
      <style>{`
        @keyframes vsm-pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
        @keyframes vsm-spin { to{transform:rotate(360deg)} }
        @keyframes vsm-fade-in { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }
        @keyframes vsm-scale-in { from{opacity:0;transform:scale(0.92)} to{opacity:1;transform:scale(1)} }
        .vsm-card { animation: vsm-scale-in 0.3s cubic-bezier(.34,1.56,.64,1) both; }
        .vsm-spinner { animation: vsm-spin 0.9s linear infinite; }
        .vsm-fade { animation: vsm-fade-in 0.4s ease both; }
        .vsm-btn:hover { opacity:0.88; transform:translateY(-1px); }
        .vsm-btn { transition: all 0.2s ease; cursor:pointer; }
      `}</style>

      <div className="vsm-card" style={{ ...CARD_STYLE, width: isMobile ? '95vw' : '100%', maxWidth: isMobile ? '95vw' : 600, maxHeight: '90vh', overflowY: 'auto' }}>
        <div style={{ height: 4, background: 'linear-gradient(90deg, #3b82f6, #8b5cf6, #06b6d4)' }} />
        <div style={{ padding: isMobile ? '18px 14px 16px' : '28px 32px 24px' }}>
          {/* Title row */}
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 24 }}>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0 }}><Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 18 }} /></div>
                <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#f1f5f9' }}>Vessel Suggestions</h2>
              </div>
              <p style={{ margin: 0, fontSize: 12, color: '#64748b' }}>
                {files.length > 0
                  ? `Analysing ${files.length} uploaded file${files.length !== 1 ? 's' : ''} to identify vessel details`
                  : 'Review detected suggestions, scan documents, or create new vessels in the Term Store'}
              </p>
            </div>
            <button className="vsm-btn" onClick={close} style={{ background: 'rgba(255,255,255,0.06)', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 8, width: 44, height: 44, color: '#94a3b8', fontSize: 18, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 18 }} /></button>
          </div>

          <StepIndicator step={step} />

          {/* STEP: EXTRACT / MATCH */}
          {(step === 'extract' || step === 'match') && (
            <div className="vsm-fade">
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 20 }}>
                {extracting ? (
                  <div className="vsm-spinner" style={{ width: 20, height: 20, border: '2px solid rgba(99,179,237,0.2)', borderTopColor: '#60a5fa', borderRadius: '50%', flexShrink: 0 }} />
                ) : (
                  <span style={{ fontSize: 18 }}><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 18 }} /></span>
                )}
                <span style={{ color: '#93c5fd', fontSize: 13, fontWeight: 600 }}>
                  {step === 'extract' ? 'Scanning file names & OCR content for vessel identity…' : 'Matching against Term Store…'}
                </span>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 20 }}>
                {files.slice(0, 6).map((f, i) => (
                  <div key={i} style={{ background: 'rgba(59,130,246,0.1)', border: '1px solid rgba(59,130,246,0.25)', borderRadius: 6, padding: '3px 10px', fontSize: 11, color: '#93c5fd', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}><Icon iconName="Page" aria-hidden="true" style={{ fontSize: 11 }} /> {f.name}</div>
                ))}
                {files.length > 6 && <div style={{ fontSize: 11, color: '#64748b', padding: '3px 4px' }}>+{files.length - 6} more</div>}
              </div>
              {step === 'match' && suggestion && (
                <div style={GLASS_PANEL}>
                  <div style={{ fontSize: 11, color: '#64748b', marginBottom: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Extracted Information</div>
                  <ExtractRow label="Vessel Name" value={suggestion.vesselName} delay={0} />
                  <ExtractRow label="Hull Number" value={suggestion.hullNumber} delay={100} />
                  <ExtractRow label="IMO Number" value={suggestion.imoNumber} delay={200} />
                  <ExtractRow label="Shipyard" value={suggestion.shipyard} delay={300} />
                  <ExtractRow label="Vessel Type" value={suggestion.vesselType} delay={400} />
                </div>
              )}
            </div>
          )}

          {/* STEP: REVIEW - NO NEW VESSELS DETECTED */}
          {step === 'review' && !suggestion && (
            <div className="vsm-fade" style={{ textAlign: 'center', padding: '32px 16px' }}>
              <div style={{ width: 64, height: 64, borderRadius: '50%', margin: '0 auto 16px', background: 'rgba(74,222,128,0.1)', border: '1px solid rgba(74,222,128,0.3)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 30 }}>
                <Icon iconName="Brightness" aria-hidden="true" style={{ fontSize: 30 }} />
              </div>
              <h3 style={{ color: '#f1f5f9', fontSize: 16, fontWeight: 700, margin: '0 0 8px' }}>
                All Documents Match Existing Vessels
              </h3>
              <p style={{ color: '#94a3b8', fontSize: 12, maxWidth: 440, margin: '0 auto 24px', lineHeight: 1.6 }}>
                The uploaded documents belong to vessels that already exist in the Term Store. No new vessels need to be created.
              </p>
              <div style={{ display: 'flex', justifyContent: 'center', gap: 12 }}>
                <label style={{
                  background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)', borderRadius: 10, padding: '10px 18px',
                  color: '#fff', fontSize: 13, fontWeight: 700, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 8,
                }}>
                  <input
                    type="file"
                    multiple
                    style={{ display: 'none' }}
                    onChange={e => {
                      const filesList = Array.from(e.target.files || []);
                      if (filesList.length === 0) return;
                      host._openVesselSuggestions(filesList);
                    }}
                  />
                  <span><Icon iconName="Page" aria-hidden="true" style={{ fontSize: 13 }} /></span> Scan Document for New Vessel
                </label>
                <button
                  type="button"
                  onClick={close}
                  style={{
                    background: 'rgba(255,255,255,0.06)', border: '1px solid rgba(255,255,255,0.1)',
                    borderRadius: 10, padding: '10px 18px', color: '#94a3b8', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                  }}
                >
                  Close
                </button>
              </div>
            </div>
          )}

          {/* STEP: REVIEW - NEW VESSEL SUGGESTION PRESENT */}
          {step === 'review' && suggestion && editedSuggestion && (
            <div className="vsm-fade">
              {/* Multi-Vessel Selector if multiple new vessels were detected */}
              {allSuggestions && allSuggestions.length > 1 && (
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 8 }}>
                    <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 11 }} /> Suggested New Vessels ({allSuggestions.length})
                  </div>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    {allSuggestions.map((s, idx) => {
                      const isSelected = idx === selectedSuggestionIndex;
                      const isOcr = !!s.ocrVerified;
                      return (
                        <button
                          key={s.vesselName + idx}
                          type="button"
                          onClick={() => {
                            host.setState(prev => ({
                              vesselSuggestionDialog: prev.vesselSuggestionDialog
                                ? {
                                    ...prev.vesselSuggestionDialog,
                                    selectedSuggestionIndex: idx,
                                    suggestion: s,
                                    editedSuggestion: {
                                      ...s,
                                      vesselName: s.vesselName,
                                      imoNumber: s.imoNumber,
                                      shipyard: s.shipyard,
                                      hullNumber: s.hullNumber,
                                      vesselType: s.vesselType,
                                    },
                                  }
                                : null,
                            }));
                          }}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 6,
                            padding: '7px 14px', borderRadius: 10,
                            background: isSelected ? 'linear-gradient(135deg, #3b82f6, #8b5cf6)' : 'rgba(255,255,255,0.06)',
                            border: isSelected ? '1px solid rgba(99,179,237,0.6)' : '1px solid rgba(255,255,255,0.1)',
                            color: isSelected ? '#fff' : '#cbd5e1',
                            fontSize: 12, fontWeight: isSelected ? 700 : 500,
                            cursor: 'pointer', transition: 'all 0.15s ease',
                            boxShadow: isSelected ? '0 4px 12px rgba(59,130,246,0.4)' : 'none',
                          }}
                        >
                          <span><Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /></span>
                          <span>{s.vesselName}</span>
                          <span style={{ fontSize: 10, opacity: 0.8 }} title={formatLastScanned(s.lastDetectedAt)}>
                            {s.lastDetectedAt ? new Date(s.lastDetectedAt).toLocaleDateString() : ''}
                          </span>
                          {isOcr && (
                            <span style={{
                              background: isSelected ? 'rgba(255,255,255,0.25)' : 'rgba(59,130,246,0.2)',
                              color: isSelected ? '#fff' : '#93c5fd',
                              padding: '1px 5px', borderRadius: 4, fontSize: 10, fontWeight: 700,
                            }}>
                              {Math.round((s.ocrConfidence || s.confidence) * 100)}% OCR
                            </span>
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Match / New Vessel Banner */}
              <div style={{
                ...GLASS_PANEL,
                marginBottom: 20,
                display: 'flex',
                alignItems: 'center',
                gap: 12,
                borderColor: hasMatch ? 'rgba(74,222,128,0.35)' : 'rgba(59,130,246,0.35)',
                background: hasMatch ? 'rgba(74,222,128,0.06)' : 'rgba(59,130,246,0.06)',
              }}>
                <div style={{ fontSize: 26 }}>{hasMatch ? <Icon iconName="BullseyeTarget" aria-hidden="true" style={{ fontSize: 26 }} /> : <Icon iconName="Search" aria-hidden="true" style={{ fontSize: 26 }} />}</div>
                <div style={{ flex: 1 }}>
                  {hasMatch ? (
                    <>
                      <div style={{ fontSize: 13, fontWeight: 700, color: '#4ade80', marginBottom: 2, display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span>Match Found — {suggestion.matchedExisting!.name}</span>
                        {suggestion.ocrVerified && (
                          <span style={{ background: 'rgba(74,222,128,0.2)', color: '#86efac', border: '1px solid rgba(74,222,128,0.4)', borderRadius: 6, padding: '1px 6px', fontSize: 10, fontWeight: 700 }}>
                            <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 10 }} /> OCR Verified
                          </span>
                        )}
                      </div>
                      <div style={{ fontSize: 11, color: '#94a3b8' }}>
                        This vessel already exists in the Term Store with {confPct}% match confidence.
                      </div>
                      <div style={{ fontSize: 10, color: '#64748b', marginTop: 4 }}>
                        Last scanned: {formatLastScanned(suggestion.lastDetectedAt)}
                      </div>
                    </>
                  ) : (
                    <>
                      <div style={{ fontSize: 13, fontWeight: 700, color: '#93c5fd', marginBottom: 2, display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span>Suggested New Vessel — {suggestion.vesselName || 'New Vessel'}</span>
                        {suggestion.ocrVerified && (
                          <span style={{ background: 'rgba(59,130,246,0.2)', color: '#93c5fd', border: '1px solid rgba(59,130,246,0.4)', borderRadius: 6, padding: '1px 6px', fontSize: 10, fontWeight: 700 }}>
                            <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 10 }} /> OCR Extracted
                          </span>
                        )}
                      </div>
                      <div style={{ fontSize: 11, color: '#94a3b8' }}>
                        Extracted from uploaded document text. Review details below and create in Term Store.
                      </div>
                      <div style={{ fontSize: 10, color: '#64748b', marginTop: 4 }}>
                        Last scanned: {formatLastScanned(suggestion.lastDetectedAt)}
                      </div>
                    </>
                  )}
                  {suggestion.matchedKeywords && suggestion.matchedKeywords.length > 0 && (
                    <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap', marginTop: 6 }}>
                      {suggestion.matchedKeywords.slice(0, 4).map((kw, i) => (
                        <span key={i} style={{ background: 'rgba(255,255,255,0.06)', color: '#93c5fd', fontSize: 10, borderRadius: 4, padding: '1px 6px' }}>
                          #{kw}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
                <div style={{ background: `${confidenceColor(suggestion.confidence)}22`, border: `1px solid ${confidenceColor(suggestion.confidence)}44`, borderRadius: 8, padding: '4px 10px', textAlign: 'center', flexShrink: 0 }}>
                  <div style={{ fontSize: 16, fontWeight: 800, color: confidenceColor(suggestion.confidence) }}>{confPct}%</div>
                  <div style={{ fontSize: 9, color: '#94a3b8', marginTop: 1 }}>{confidenceLabel(suggestion.confidence)}</div>
                </div>
              </div>

              <div style={{ ...GLASS_PANEL, padding: '20px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
                  <div style={{ fontSize: 11, color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}><Icon iconName="Edit" aria-hidden="true" style={{ fontSize: 11 }} /> Review &amp; Edit Vessel Details</div>
                  <label style={{
                    background: 'rgba(99,179,237,0.15)', border: '1px solid rgba(99,179,237,0.3)',
                    borderRadius: 8, padding: '4px 10px', fontSize: 11, color: '#93c5fd',
                    cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6,
                    fontWeight: 600,
                  }}>
                    <input
                      type="file"
                      multiple
                      style={{ display: 'none' }}
                      onChange={e => {
                        const filesList = Array.from(e.target.files || []);
                        if (filesList.length === 0) return;
                        host._openVesselSuggestions(filesList);
                      }}
                    />
                    <span><Icon iconName="Page" aria-hidden="true" style={{ fontSize: 11 }} /></span> Scan Document
                  </label>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0 16px' }}>
                  <Field label="Vessel Name" value={editedSuggestion.vesselName} onChange={v => setEdited({ vesselName: v })} required hint="Full display name of the vessel" />
                  <Field label="IMO Number" value={editedSuggestion.imoNumber} onChange={v => setEdited({ imoNumber: v })} required hint="7-digit IMO registration number" />
                  <Field label="Hull Number" value={editedSuggestion.hullNumber} onChange={v => setEdited({ hullNumber: v })} hint="Yard / hull build number (e.g. SS268)" />
                  <Field label="Vessel Type" value={editedSuggestion.vesselType} onChange={v => setEdited({ vesselType: v })} type="select" options={VESSEL_TYPES} />
                  <div style={{ gridColumn: '1 / -1' }}>
                    <Field label="Shipyard Name" value={editedSuggestion.shipyard} onChange={v => setEdited({ shipyard: v })} hint="Name of the shipyard that built the vessel" />
                  </div>
                </div>
                <div style={{ marginTop: 4 }}>
                  <div style={{ fontSize: 10, color: '#475569', marginBottom: 6, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Extracted from</div>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                    {(editedSuggestion.sourceFiles || []).slice(0, 4).map((f, i) => (
                      <div key={i} style={{ background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.06)', borderRadius: 4, padding: '2px 8px', fontSize: 10, color: '#64748b', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f}</div>
                    ))}
                  </div>
                </div>
              </div>

              {error && (
                <div style={{ marginTop: 12, padding: '10px 14px', background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 10, fontSize: 13, color: '#fca5a5' }}>
                  <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 13 }} /> {error}
                </div>
              )}

              <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
                {hasMatch && (
                  <button id="vsm-use-existing-btn" className="vsm-btn" onClick={useExisting} disabled={creating} style={{ flex: 1, padding: '13px 20px', borderRadius: 12, background: creating ? 'rgba(16,185,129,0.2)' : 'linear-gradient(135deg, #059669, #10b981)', border: creating ? '1px solid rgba(16,185,129,0.35)' : 'none', color: '#fff', fontSize: 14, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, opacity: creating ? 0.85 : 1 }}>
                    <span style={{ fontSize: 18 }}>{creating ? <Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 18 }} /> : <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 18 }} />}</span>{creating ? 'Routing to Matched Vessel…' : 'Use Existing Vessel'}
                  </button>
                )}
                <button
                  id="vsm-create-vessel-btn"
                  className="vsm-btn"
                  onClick={() => void host._applyVesselSuggestionCreate()}
                  disabled={creating || !editedSuggestion.vesselName.trim()}
                  style={{ flex: 1, padding: '13px 20px', borderRadius: 12, background: creating ? 'rgba(99,179,237,0.15)' : 'linear-gradient(135deg, #3b82f6, #8b5cf6)', border: creating ? '1px solid rgba(99,179,237,0.3)' : 'none', color: creating ? '#93c5fd' : '#fff', fontSize: 14, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, opacity: !editedSuggestion.vesselName.trim() ? 0.5 : 1 }}
                >
                  {creating ? (
                    <><div className="vsm-spinner" style={{ width: 16, height: 16, border: '2px solid rgba(147,197,253,0.3)', borderTopColor: '#93c5fd', borderRadius: '50%' }} />Creating in Term Store…</>
                  ) : (
                    <><span style={{ fontSize: 18 }}><Icon iconName="Add" aria-hidden="true" style={{ fontSize: 18 }} /></span>{hasMatch ? 'Re-create / Update Vessel' : 'Create New Vessel'}</>
                  )}
                </button>
                <button
                  id="vsm-ignore-vessel-btn"
                  className="vsm-btn"
                  onClick={() => host._ignoreCurrentVesselSuggestion()}
                  disabled={creating}
                  style={{ padding: '13px 16px', borderRadius: 12, background: 'rgba(239,68,68,0.08)', border: '1px solid rgba(239,68,68,0.35)', color: '#fca5a5', fontSize: 13, minWidth: 108 }}
                  title="Mark this suggestion as false positive"
                >
                  Ignore
                </button>
                <button className="vsm-btn" onClick={close} style={{ padding: '13px 16px', borderRadius: 12, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: '#94a3b8', fontSize: 13 }}>Skip</button>
              </div>
            </div>
          )}

          {/* STEP: DONE */}
          {step === 'done' && (
            <div className="vsm-fade" style={{ textAlign: 'center', padding: '16px 0 8px' }}>
              <div style={{ width: 72, height: 72, borderRadius: '50%', margin: '0 auto 20px', background: 'linear-gradient(135deg, #059669, #10b981)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 34, boxShadow: '0 0 32px rgba(16,185,129,0.4)' }}>
                {editedSuggestion?.matchedExisting ? <Icon iconName="BullseyeTarget" aria-hidden="true" style={{ fontSize: 34 }} /> : <Icon iconName="Completed" aria-hidden="true" style={{ fontSize: 34 }} />}
              </div>
              <h3 style={{ color: '#f1f5f9', fontSize: 18, fontWeight: 700, margin: '0 0 8px' }}>
                {editedSuggestion?.matchedExisting ? 'Linked to Existing Vessel' : 'Vessel Created Successfully!'}
              </h3>
              <p style={{ color: '#94a3b8', fontSize: 13, margin: '0 0 24px', lineHeight: 1.6 }}>
                {editedSuggestion?.matchedExisting
                  ? `Documents will be tagged under "${editedSuggestion.matchedExisting.name}".`
                  : `"${editedSuggestion?.vesselName || 'Vessel'}" has been added to the Term Store and folder provisioning is underway.`
                }
              </p>
              {editedSuggestion && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, justifyContent: 'center', marginBottom: 24 }}>
                  {[
                    { icon: <Icon iconName="NumberSymbol" aria-hidden="true" style={{ fontSize: 12 }} />, label: 'IMO', value: editedSuggestion.imoNumber },
                    { icon: <Icon iconName="ConstructionCone" aria-hidden="true" style={{ fontSize: 12 }} />, label: 'Hull', value: editedSuggestion.hullNumber },
                    { icon: <Icon iconName="Manufacturing" aria-hidden="true" style={{ fontSize: 12 }} />, label: 'Shipyard', value: editedSuggestion.shipyard },
                    { icon: <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} />, label: 'Type', value: editedSuggestion.vesselType },
                  ].filter(p => p.value).map((p, i) => (
                    <div key={i} style={{ background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 8, padding: '6px 12px', fontSize: 12 }}>
                      <span style={{ marginRight: 6 }}>{p.icon}</span>
                      <span style={{ color: '#64748b', marginRight: 4 }}>{p.label}:</span>
                      <span style={{ color: '#e2e8f0', fontWeight: 600 }}>{p.value}</span>
                    </div>
                  ))}
                </div>
              )}
              <button id="vsm-done-btn" className="vsm-btn" onClick={close} style={{ padding: '12px 32px', borderRadius: 12, background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)', border: 'none', color: '#fff', fontSize: 14, fontWeight: 700 }}>Done</button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
