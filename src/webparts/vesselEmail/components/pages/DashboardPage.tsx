import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge, cleanName } from '../constants';
import { resolveImgUrl } from '../vesselImagePool';

export interface RealDashboardDoc {
  id: string;
  name: string;
  vessel: string;
  type: string;
  modified: string;
  modifiedEpoch: number;
  status: 'Valid' | 'Expiring Soon' | 'Expired';
  fileSize?: string;
  subFolderPath?: string;
  webUrl?: string;
}

export interface DashboardStats {
  total_documents: number;
  total_vessels: number;
  pending_approvals: number;
  expiring_soon_count: number;
  expired_count: number;
  valid_count: number;
  recent_documents: RealDashboardDoc[];
  documents?: RealDashboardDoc[];
  expiry_overview: {
    total: number;
    expired: number;
    expiring_soon: number;
    valid: number;
    expired_pct: number;
    expiring_pct: number;
    valid_pct: number;
  };
}

/**
 * Extracts all real unique documents across rows, uploadedFilesByFolder,
 * live SharePoint drive nodes, and backend dashboard stats.
 */
export function extractRealDocuments(host: VesselEmail): RealDashboardDoc[] {
  const { rows, uploadedFilesByFolder, vessels, dashboardStats } = host.state;
  const seenKeys = new Set<string>();
  const realDocs: RealDashboardDoc[] = [];

  const now = Date.now();
  const thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
  const oneYearMs = 365 * 24 * 60 * 60 * 1000;

  const addDoc = (doc: RealDashboardDoc): void => {
    if (!doc.name) return;
    const key = `${(doc.name || '').trim().toLowerCase()}||${(doc.vessel || '').trim().toLowerCase()}||${(doc.subFolderPath || '').trim().toLowerCase()}`;
    if (seenKeys.has(key)) return;
    seenKeys.add(key);
    realDocs.push(doc);
  };

  // 0. Include backend dashboard documents if available (from fast Graph search)
  const dashDocs = dashboardStats?.documents || dashboardStats?.recent_documents;
  if (Array.isArray(dashDocs) && dashDocs.length > 0) {
    dashDocs.forEach(doc => {
      if (doc && doc.name) addDoc(doc);
    });
  }

  // 1. Scan rows from backend database / flat-tree
  (rows || []).forEach(r => {
    if (!r.fileName) return;
    const modEpoch = r.fileUploadedAt
      ? (typeof r.fileUploadedAt === 'number' ? r.fileUploadedAt : Date.parse(r.fileUploadedAt as any) || now)
      : now;
    const dateStr = r.fileUploadedAt
      ? new Date(modEpoch).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
      : 'Today';

    const typeLabel = r.subCategory || r.category || r.group || 'Document';


    // Compliance / Certificate expiry evaluation
    const isCertOrInsurance = /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(r.category || '') ||
                             /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(r.subCategory || '') ||
                             /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(r.fileName || '');

    let status: 'Valid' | 'Expiring Soon' | 'Expired' = 'Valid';
    if (isCertOrInsurance) {
      const expiryEpoch = modEpoch + oneYearMs;
      if (expiryEpoch < now) {
        status = 'Expired';
      } else if (expiryEpoch <= now + thirtyDaysMs) {
        status = 'Expiring Soon';
      } else {
        status = 'Valid';
      }
    }

    addDoc({
      id: r.fileId || r.fileName,
      name: r.fileName,
      vessel: r.vesselName && r.vesselName !== 'Common for all vessels' && r.vesselName !== 'Kaizen - Knowledge Bank'
        ? r.vesselName
        : (r.group || 'Shared Documents'),
      type: typeLabel,
      modified: dateStr,
      modifiedEpoch: modEpoch,
      status,
      fileSize: r.fileSize || '—',
      subFolderPath: r.subFolderPath,
    });
  });

  // 2. Scan uploadedFilesByFolder
  Object.entries(uploadedFilesByFolder || {}).forEach(([folderKey, list]) => {
    if (!Array.isArray(list)) return;
    list.forEach(f => {
      if (!f?.name) return;
      const modEpoch = f.uploadedAt || now;
      const dateStr = f.date || new Date(modEpoch).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
      const matchedVessel = (vessels || []).find(v => folderKey.toLowerCase().includes(v.name.toLowerCase()));
      const isCertOrInsurance = /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(folderKey) ||
                               /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(f.name);
      let status: 'Valid' | 'Expiring Soon' | 'Expired' = 'Valid';
      if (isCertOrInsurance) {
        const expiryEpoch = modEpoch + oneYearMs;
        if (expiryEpoch < now) {
          status = 'Expired';
        } else if (expiryEpoch <= now + thirtyDaysMs) {
          status = 'Expiring Soon';
        } else {
          status = 'Valid';
        }
      }

      addDoc({
        id: (f as any).id || f.name,
        name: f.name,
        vessel: matchedVessel ? matchedVessel.name : 'Shared Documents',
        type: 'Uploaded Document',
        modified: dateStr,
        modifiedEpoch: modEpoch,
        status,
        fileSize: f.size || '—',
        subFolderPath: folderKey,
      });
    });
  });

  // 3. Scan live spoFolderMap
  for (const [, node] of Array.from(host.state.spoFolderMap.entries())) {
    if (!node || node.deleted || (node.id && host._appDeletedItemIds.has(node.id))) continue;
    (node.children || []).forEach(child => {
      if (!child.isFolder && child.name && (!child.id || !host._appDeletedItemIds.has(child.id))) {
        const createdStr = (child as any).createdDateTime || (child as any).lastModifiedDateTime;
        const modEpoch = createdStr ? Date.parse(createdStr) || now : now;
        const dateStr = createdStr ? new Date(modEpoch).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : 'Today';
        const matchedVessel = (vessels || []).find(v => (node.serverRelativePath || node.name || '').toLowerCase().includes(v.name.toLowerCase()));
        const isCertOrInsurance = /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(node.name || '') ||
                                 /insurance|certificate|survey|crew|audit|statutory|compliance/i.test(child.name);
        let status: 'Valid' | 'Expiring Soon' | 'Expired' = 'Valid';
        if (isCertOrInsurance) {
          const expiryEpoch = modEpoch + oneYearMs;
          if (expiryEpoch < now) {
            status = 'Expired';
          } else if (expiryEpoch <= now + thirtyDaysMs) {
            status = 'Expiring Soon';
          } else {
            status = 'Valid';
          }
        }

        addDoc({
          id: child.id || child.name,
          name: child.name,
          vessel: matchedVessel ? matchedVessel.name : (node.name || 'Shared Documents'),
          type: node.name || 'Document',
          modified: dateStr,
          modifiedEpoch: modEpoch,
          status,
          fileSize: typeof child.size === 'number' ? `${(child.size / 1024).toFixed(1)} KB` : '—',
          subFolderPath: node.serverRelativePath || node.name,
        });
      }
    });
  }

  return realDocs;
}

function getFileIcon(name: string): { icon: string; color: string } {
  const low = (name || '').toLowerCase();
  if (low.endsWith('.pdf')) return { icon: '📄', color: '#ef4444' };
  if (low.endsWith('.docx') || low.endsWith('.doc')) return { icon: '📝', color: '#2563eb' };
  if (low.endsWith('.xlsx') || low.endsWith('.xls') || low.endsWith('.csv')) return { icon: '📊', color: '#10b981' };
  if (low.endsWith('.dwg') || low.endsWith('.dxf')) return { icon: '📐', color: '#8b5cf6' };
  if (low.endsWith('.png') || low.endsWith('.jpg') || low.endsWith('.jpeg')) return { icon: '🖼️', color: '#f59e0b' };
  return { icon: '📄', color: '#64748b' };
}

function getTypeBadgeColor(type: string): 'blue' | 'green' | 'orange' | 'purple' | 'red' | 'default' {
  const low = (type || '').toLowerCase();
  if (low.includes('insurance') || low.includes('claim')) return 'orange';
  if (low.includes('crew')) return 'blue';
  if (low.includes('tech') || low.includes('maint')) return 'green';
  if (low.includes('cert') || low.includes('statutory')) return 'purple';
  if (low.includes('incident') || low.includes('danger')) return 'red';
  if (low.includes('drawing')) return 'blue';
  if (low.includes('manual')) return 'green';
  return 'default';
}

function DashboardCount({ value, isLoading }: { value: number; isLoading: boolean }): React.ReactElement {
  const [displayValue, setDisplayValue] = React.useState(0);

  React.useEffect(() => {
    if (isLoading || value <= 0) {
      setDisplayValue(0);
      return undefined;
    }

    const duration = Math.min(1000, Math.max(350, value * 14));
    const startedAt = performance.now();
    let animationFrame = 0;

    const tick = (now: number): void => {
      const progress = Math.min((now - startedAt) / duration, 1);
      setDisplayValue(Math.round(value * (1 - Math.pow(1 - progress, 3))));
      if (progress < 1) animationFrame = requestAnimationFrame(tick);
    };

    animationFrame = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(animationFrame);
  }, [isLoading, value]);

  return <>{displayValue}</>;
}

export function renderDashboard(host: VesselEmail): React.ReactElement {
  const { vessels, approvalsList, loading, dashboardStats } = host.state;
  const userDisplayName = host.props.userDisplayName || (host.props.userEmail ? host.props.userEmail.split('@')[0] : 'SPE ADMIN');
  const isDashboardLoading = loading && !dashboardStats;

  // Extract all real documents from current live state
  const realDocs = extractRealDocuments(host);
  const totalDocs = dashboardStats?.total_documents ?? realDocs.length;
  const expiredDocs = dashboardStats?.expired_count ?? realDocs.filter(d => d.status === 'Expired').length;
  const expiringSoonDocs = dashboardStats?.expiring_soon_count ?? realDocs.filter(d => d.status === 'Expiring Soon').length;
  const validDocs = dashboardStats?.valid_count ?? realDocs.filter(d => d.status === 'Valid').length;
  const pendingApprovals = dashboardStats?.pending_approvals ?? (approvalsList || []).filter(a => a.status === 'Pending').length;
  const totalVessels = dashboardStats?.total_vessels ?? (vessels || []).length;

  // Recent documents sorted by newest modifiedEpoch
  const recentDocs = (dashboardStats?.recent_documents && dashboardStats.recent_documents.length > 0)
    ? dashboardStats.recent_documents.slice(0, 6)
    : realDocs
        .slice()
        .sort((a, b) => b.modifiedEpoch - a.modifiedEpoch)
        .slice(0, 6);

  // Dynamic percentages for Donut Chart
  const expiredPct = dashboardStats?.expiry_overview?.expired_pct ?? (totalDocs > 0 ? Math.round((expiredDocs / totalDocs) * 100) : 0);
  const expiringPct = dashboardStats?.expiry_overview?.expiring_pct ?? (totalDocs > 0 ? Math.round((expiringSoonDocs / totalDocs) * 100) : 0);
  const validPct = dashboardStats?.expiry_overview?.valid_pct ?? (totalDocs > 0 ? Math.max(0, 100 - expiredPct - expiringPct) : 0);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* Welcome Banner */}
      <div style={{
        background: 'linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)',
        borderRadius: 12, padding: '24px 32px', display: 'flex', justifyContent: 'space-between',
        alignItems: 'center', border: '1px solid #bfdbfe', position: 'relative', overflow: 'hidden',
        boxShadow: '0 1px 4px rgba(0,0,0,0.04)',
      }}>
        <div>
          <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#1e3a8a' }}>
            Welcome back, {userDisplayName}! 👋
          </h1>
          <p style={{ margin: '6px 0 0', fontSize: 13, color: '#3b82f6' }}>
            Here's the live overview of your fleet vessels and managed documents.
          </p>
        </div>
        {/* Vessel Graphic Banner - Animated Sailing Ship */}
        <div style={{
          height: 88,
          width: 215,
          borderRadius: 8,
          overflow: 'hidden',
          flexShrink: 0,
          boxShadow: '0 4px 14px rgba(37,99,235,0.18)',
          border: '1px solid #bfdbfe',
          background: 'linear-gradient(180deg, #b3d4f0 0%, #0a3d72 100%)',
          position: 'relative',
        }}>
          <AnimatedVesselBanner />
        </div>
      </div>

        {/* 4 Stat Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16 }}>
          {/* Stat 1: Total Documents */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a', minHeight: 34 }}><DashboardCount value={totalDocs} isLoading={isDashboardLoading} /></div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Total Documents</div>
              <button onClick={() => host._goToView('list')} style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#f0f9ff', color: '#0284c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>📄</div>
          </div>

          {/* Stat 2: Expiring Soon */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: expiringSoonDocs > 0 ? '#f59e0b' : '#0f172a', minHeight: 34 }}><DashboardCount value={expiringSoonDocs} isLoading={isDashboardLoading} /></div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Expiring Soon</div>
              <button onClick={() => host._goToView('reports')} style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: expiringSoonDocs > 0 ? '#fffbeb' : '#ecfdf5', color: expiringSoonDocs > 0 ? '#f59e0b' : '#10b981', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>
              {expiringSoonDocs > 0 ? '⚠️' : '✅'}
            </div>
          </div>

          {/* Stat 3: Pending Approvals */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: pendingApprovals > 0 ? '#ea580c' : '#0f172a', minHeight: 34 }}>
                <DashboardCount value={pendingApprovals} isLoading={isDashboardLoading} />
              </div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Pending Approvals</div>
              <button onClick={() => host._goToView('approvals')} style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: pendingApprovals > 0 ? '#fff7ed' : '#f8fafc', color: pendingApprovals > 0 ? '#ea580c' : '#64748b', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>
              {pendingApprovals > 0 ? '⏰' : '✓'}
            </div>
          </div>

          {/* Stat 4: Total Vessels */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a', minHeight: 34 }}><DashboardCount value={totalVessels} isLoading={isDashboardLoading} /></div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Total Vessels</div>
              <button onClick={() => host._goToView('vessels')} style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#eff6ff', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>🚢</div>
          </div>
        </div>

        {/* Main Dashboard Grid: Recent Docs + Expiry Overview */}
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 20 }}>
          {/* Recent Documents Table */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20, boxShadow: '0 1px 4px rgba(0,0,0,0.03)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Recent Documents</h3>
                {totalDocs > 0 && (
                  <span style={{ fontSize: 11, fontWeight: 700, color: '#0284c7', background: '#e0f2fe', borderRadius: 12, padding: '1px 8px' }}>
                    {totalDocs} total
                  </span>
                )}
              </div>
              <button onClick={() => host._goToView('list')} style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>View all</button>
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: '2px solid #e2e8f0', textTransform: 'uppercase', fontSize: 11, color: '#64748b', textAlign: 'left', letterSpacing: '0.04em' }}>
                  <th style={{ padding: '8px 12px' }}>Name</th>
                  <th style={{ padding: '8px 12px' }}>Vessel</th>
                  <th style={{ padding: '8px 12px' }}>Type</th>
                  <th style={{ padding: '8px 12px' }}>Modified</th>
                </tr>
              </thead>
              <tbody>
                {recentDocs.length === 0 ? (
                  <tr>
                    <td colSpan={4} style={{ padding: '36px 16px', textAlign: 'center', color: '#94a3b8' }}>
                      <div style={{ fontSize: 28, marginBottom: 6 }}>📂</div>
                      <div style={{ fontWeight: 600, color: '#64748b' }}>No documents uploaded yet</div>
                      <div style={{ fontSize: 12, marginTop: 4, color: '#94a3b8' }}>Uploaded documents across vessels and departments will appear here automatically.</div>
                    </td>
                  </tr>
                ) : (
                  recentDocs.map(doc => {
                    const { icon, color } = getFileIcon(doc.name);
                    return (
                      <tr
                        key={doc.id}
                        style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.12s' }}
                        onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                        onMouseLeave={e => (e.currentTarget.style.background = '#fff')}
                      >
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b', maxWidth: 220 }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <span style={{ fontSize: 16, color, flexShrink: 0 }}>{icon}</span>
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={doc.name}>
                              {doc.name}
                            </span>
                          </div>
                        </td>
                        <td style={{ padding: '10px 12px', color: '#475569', whiteSpace: 'nowrap' }}>
                          <span style={{
                            display: 'inline-flex', alignItems: 'center', gap: 4,
                            background: '#f1f5f9', color: '#334155', borderRadius: 6,
                            padding: '2px 8px', fontSize: 11, fontWeight: 600,
                          }}>
                            {doc.vessel}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', whiteSpace: 'nowrap' }}>
                          {badge(getTypeBadgeColor(doc.type), doc.type)}
                        </td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 12, whiteSpace: 'nowrap' }}>
                          {doc.modified}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Document Expiry Overview Chart */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20, boxShadow: '0 1px 4px rgba(0,0,0,0.03)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Document Expiry Overview</h3>
              <button onClick={() => host._goToView('reports')} style={{ background: 'none', border: 'none', color: '#0284c7', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>View all</button>
            </div>

            {/* SVG Donut Chart */}
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16 }}>
              <div style={{ position: 'relative', width: 140, height: 140 }}>
                <svg width="140" height="140" viewBox="0 0 42 42">
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f1f5f9" strokeWidth="5" />
                  {totalDocs > 0 ? (
                    <>
                      {/* Valid (Green) */}
                      {validPct > 0 && (
                        <circle
                          cx="21" cy="21" r="15.91549430918954" fill="transparent"
                          stroke="#10b981" strokeWidth="5"
                          strokeDasharray={`${validPct} ${100 - validPct}`}
                          strokeDashoffset="25"
                        />
                      )}
                      {/* Expiring in 30 days (Orange) */}
                      {expiringPct > 0 && (
                        <circle
                          cx="21" cy="21" r="15.91549430918954" fill="transparent"
                          stroke="#f59e0b" strokeWidth="5"
                          strokeDasharray={`${expiringPct} ${100 - expiringPct}`}
                          strokeDashoffset={`${25 - validPct}`}
                        />
                      )}
                      {/* Expired (Red) */}
                      {expiredPct > 0 && (
                        <circle
                          cx="21" cy="21" r="15.91549430918954" fill="transparent"
                          stroke="#ef4444" strokeWidth="5"
                          strokeDasharray={`${expiredPct} ${100 - expiredPct}`}
                          strokeDashoffset={`${25 - validPct - expiringPct}`}
                        />
                      )}
                    </>
                  ) : (
                    <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#e2e8f0" strokeWidth="5" strokeDasharray="100 0" strokeDashoffset="25" />
                  )}
                </svg>
                <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
                  <span style={{ fontSize: 22, fontWeight: 800, color: '#0f172a' }}>{totalDocs}</span>
                  <span style={{ fontSize: 10, color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Total</span>
                </div>
              </div>

              {/* Legend */}
              <div style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: 8, fontSize: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#ef4444' }} />
                    <span style={{ color: '#475569' }}>Expired</span>
                  </div>
                  <span style={{ fontWeight: 700, color: '#0f172a' }}>
                    {expiredDocs} ({expiredPct}%)
                  </span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#f59e0b' }} />
                    <span style={{ color: '#475569' }}>Expiring in 30 days</span>
                  </div>
                  <span style={{ fontWeight: 700, color: '#0f172a' }}>
                    {expiringSoonDocs} ({expiringPct}%)
                  </span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#10b981' }} />
                    <span style={{ color: '#475569' }}>Valid</span>
                  </div>
                  <span style={{ fontWeight: 700, color: '#0f172a' }}>
                    {validDocs} ({validPct}%)
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
}

/**
 * AnimatedVesselBanner renders a high-definition, multi-layered SVG of a sailing
 * container ship on rolling ocean waves with drifting clouds and sun glow.
 * All layers are animated via GPU-accelerated CSS keyframes and support prefers-reduced-motion.
 */
export function AnimatedVesselBanner(): React.ReactElement {
  const containerColors = [
    '#e53935', '#1e88e5', '#43a047', '#fb8c00', '#8e24aa',
    '#039be5', '#c0ca33', '#fbc02d', '#00acc1', '#3949ab'
  ];

  // Pattern of container stack heights (14 columns)
  const stackHeights = [2, 1, 3, 2, 1, 2, 3, 2, 3, 1, 2, 3, 2, 1];

  return (
    <svg
      viewBox="0 0 400 180"
      preserveAspectRatio="xMidYMid meet"
      style={{ width: '100%', height: '100%', display: 'block' }}
      aria-label="Animated Sailing Ship Illustration"
    >
      <style>{`
        @keyframes dashShipMotion {
          0% { transform: translate(0px, 0px) rotate(0deg); }
          25% { transform: translate(2px, -3px) rotate(-1.1deg); }
          50% { transform: translate(4px, 0px) rotate(0.4deg); }
          75% { transform: translate(2px, 3px) rotate(1.1deg); }
          100% { transform: translate(0px, 0px) rotate(0deg); }
        }
        @keyframes dashWaveScroll {
          0% { transform: translateX(0px); }
          100% { transform: translateX(-400px); }
        }
        @keyframes dashCloudDrift {
          0% { transform: translateX(0px); }
          50% { transform: translateX(18px); }
          100% { transform: translateX(0px); }
        }
        @keyframes dashSprayPulse {
          0%, 100% { opacity: 0.45; transform: scale(0.95); }
          25% { opacity: 0.85; transform: scale(1.08); }
          75% { opacity: 0.35; transform: scale(0.9); }
        }
        @keyframes dashboardStatLoading {
          0%, 100% { opacity: 0.35; }
          50% { opacity: 1; }
        }
        .dashboard-stat-loading {
          display: inline-block;
          animation: dashboardStatLoading 1.1s ease-in-out infinite;
        }
        .vessel-ship-animated {
          animation: dashShipMotion 3.8s ease-in-out infinite;
          transform-origin: 200px 115px;
          transform-box: fill-box;
          will-change: transform;
        }
        .vessel-wave-back {
          animation: dashWaveScroll 6.5s linear infinite;
          will-change: transform;
        }
        .vessel-wave-mid {
          animation: dashWaveScroll 4s linear infinite;
          will-change: transform;
        }
        .vessel-wave-front {
          animation: dashWaveScroll 3s linear infinite;
          will-change: transform;
        }
        .vessel-cloud-slow {
          animation: dashCloudDrift 18s ease-in-out infinite;
          will-change: transform;
        }
        .vessel-cloud-fast {
          animation: dashCloudDrift 12s ease-in-out infinite reverse;
          will-change: transform;
        }
        .vessel-bow-spray {
          animation: dashSprayPulse 3.8s ease-in-out infinite;
          transform-origin: 390px 125px;
          transform-box: fill-box;
        }
        @media (prefers-reduced-motion: reduce) {
          .vessel-ship-animated,
          .vessel-wave-back,
          .vessel-wave-mid,
          .vessel-wave-front,
          .vessel-cloud-slow,
          .vessel-cloud-fast,
          .vessel-bow-spray,
          .dashboard-stat-loading {
            animation: none !important;
            transform: none !important;
          }
        }
      `}</style>
      <defs>
        <linearGradient id="vessel_sky" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#b3d4f0" />
          <stop offset="100%" stopColor="#dceefa" />
        </linearGradient>
        <linearGradient id="vessel_water" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#1565a8" />
          <stop offset="100%" stopColor="#0a3d72" />
        </linearGradient>
        <linearGradient id="vessel_hull" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#1a365d" />
          <stop offset="100%" stopColor="#0f2942" />
        </linearGradient>
      </defs>

      {/* ── Sky & Sun Layer ── */}
      <rect x="0" y="0" width="400" height="116" fill="url(#vessel_sky)" />

      {/* Sun with warm glow */}
      <circle cx="360" cy="30" r="22" fill="#fff9e0" opacity="0.6" />
      <circle cx="360" cy="30" r="14" fill="#ffe082" opacity="0.75" />

      {/* Drifting Clouds */}
      <g className="vessel-cloud-slow">
        <ellipse cx="60" cy="24" rx="30" ry="9" fill="#ffffff" opacity="0.75" />
        <ellipse cx="80" cy="20" rx="18" ry="7" fill="#ffffff" opacity="0.65" />
        <ellipse cx="290" cy="26" rx="32" ry="8" fill="#ffffff" opacity="0.6" />
      </g>
      <g className="vessel-cloud-fast">
        <ellipse cx="180" cy="18" rx="26" ry="8" fill="#ffffff" opacity="0.65" />
        <ellipse cx="198" cy="14" rx="16" ry="6" fill="#ffffff" opacity="0.55" />
      </g>

      {/* Horizon haze */}
      <rect x="0" y="110" width="400" height="8" fill="#1565a8" opacity="0.3" />

      {/* ── Ocean Water Base ── */}
      <rect x="0" y="116" width="400" height="64" fill="url(#vessel_water)" />

      {/* ── Background Waves (Behind Ship) ── */}
      <g className="vessel-wave-back">
        <path
          d="M0,118 Q25,115 50,118 Q75,121 100,118 Q125,115 150,118 Q175,121 200,118 Q225,115 250,118 Q275,121 300,118 Q325,115 350,118 Q375,121 400,118 Q425,115 450,118 Q475,121 500,118 Q525,115 550,118 Q575,121 600,118 Q625,115 650,118 Q675,121 700,118 Q725,115 750,118 Q775,121 800,118"
          stroke="#ffffff"
          strokeWidth="0.8"
          fill="none"
          opacity="0.35"
        />
      </g>

      {/* ── Animated Sailing Ship Group ── */}
      <g className="vessel-ship-animated">
        {/* Deck Containers */}
        {stackHeights.map((layers, col) => {
          const cx = 40 + col * 18;
          return (
            <g key={col}>
              {Array.from({ length: layers }).map((_, layer) => {
                const cy = 92 - layer * 13;
                const cc = containerColors[(col * 2 + layer) % containerColors.length];
                return (
                  <g key={layer}>
                    <rect
                      x={cx}
                      y={cy}
                      width="16"
                      height="12"
                      fill={cc}
                      rx="0.8"
                      stroke="#111"
                      strokeWidth="0.4"
                    />
                    <line
                      x1={cx + 2}
                      y1={cy + 3}
                      x2={cx + 14}
                      y2={cy + 3}
                      stroke="rgba(0,0,0,0.2)"
                      strokeWidth="0.4"
                    />
                    <line
                      x1={cx + 2}
                      y1={cy + 7}
                      x2={cx + 14}
                      y2={cy + 7}
                      stroke="rgba(0,0,0,0.2)"
                      strokeWidth="0.4"
                    />
                  </g>
                );
              })}
            </g>
          );
        })}

        {/* Aft Superstructure / Bridge */}
        <g>
          {/* Main White Tower */}
          <rect x="296" y="56" width="38" height="48" fill="#f8fafc" rx="2" stroke="#c0c8d0" strokeWidth="0.5" />
          {/* Bridge Windows */}
          <rect x="298" y="60" width="34" height="5" fill="#64b5f6" />
          {/* Cabin Windows */}
          <rect x="300" y="68" width="7" height="6" fill="#90a4ae" rx="1" />
          <rect x="311" y="68" width="7" height="6" fill="#90a4ae" rx="1" />
          <rect x="322" y="68" width="7" height="6" fill="#90a4ae" rx="1" />
          <rect x="300" y="78" width="7" height="6" fill="#90a4ae" rx="1" />
          <rect x="311" y="78" width="7" height="6" fill="#90a4ae" rx="1" />
          <rect x="322" y="78" width="7" height="6" fill="#90a4ae" rx="1" />
          {/* Radar Mast & Antenna */}
          <line x1="315" y1="36" x2="315" y2="56" stroke="#94a3b8" strokeWidth="1.5" />
          <line x1="309" y1="42" x2="321" y2="42" stroke="#94a3b8" strokeWidth="1" />
          <circle cx="315" cy="35" r="2" fill="#ef4444" />
          {/* Funnel / Smokestack */}
          <rect x="318" y="32" width="14" height="24" fill="#dd6b20" rx="2" />
          <rect x="321" y="26" width="8" height="6" fill="#1a365d" rx="1" />
          <rect x="316" y="30" width="18" height="2" fill="#1a365d" />
        </g>

        {/* Main Hull */}
        <path
          d="M 22,102 L 358,102 Q 376,102 383,112 L 388,126 L 14,126 Q 10,126 15,116 Z"
          fill="url(#vessel_hull)"
        />

        {/* Deck Rail Top */}
        <rect x="22" y="100" width="338" height="4" fill="#243650" rx="1" />

        {/* Bow Accent */}
        <path d="M 358,102 L 390,114 L 386,126 L 358,126 Z" fill="#1a365d" />
        <path d="M 380,104 L 394,118 L 392,126 L 386,126 L 390,114 Z" fill="#0f2942" />

        {/* Stern Curve */}
        <path d="M 22,102 L 12,112 L 14,126 L 22,126 Z" fill="#0f2942" />

        {/* Waterline / Red Keel Stripe */}
        <rect x="14" y="120" width="374" height="6" fill="#c53030" />

        {/* Portholes */}
        {Array.from({ length: 15 }).map((_, idx) => (
          <circle key={idx} cx={42 + idx * 16} cy={114} r="2.2" fill="#fff" opacity="0.75" />
        ))}

        {/* Hull Shadow Band */}
        <rect x="14" y="117" width="374" height="3" fill="rgba(0,0,0,0.18)" />

        {/* Bow Wave Spray & Wake */}
        <g className="vessel-bow-spray">
          <path d="M 388,122 Q 396,128 392,136 Q 387,142 378,138" stroke="#ffffff" strokeWidth="1.6" fill="none" opacity="0.6" />
          <circle cx="394" cy="126" r="1.5" fill="#ffffff" opacity="0.7" />
          <circle cx="397" cy="130" r="1.2" fill="#ffffff" opacity="0.6" />
        </g>
        {/* Stern Wake */}
        <path d="M 14,122 Q 6,128 10,136 Q 16,142 24,138" stroke="#ffffff" strokeWidth="1.5" fill="none" opacity="0.6" />
      </g>

      {/* ── Midground Rolling Wave Ripples (In Front of Hull) ── */}
      <g className="vessel-wave-mid">
        <path
          d="M0,128 Q40,124 80,128 Q120,132 160,128 Q200,124 240,128 Q280,132 320,128 Q360,124 400,128 Q440,124 480,128 Q520,132 560,128 Q600,124 640,128 Q680,132 720,128 Q760,124 800,128"
          stroke="#ffffff"
          strokeWidth="1.2"
          fill="none"
          opacity="0.45"
        />
      </g>

      {/* ── Foreground Shimmering Waves ── */}
      <g className="vessel-wave-front">
        <path
          d="M0,138 Q50,134 100,138 Q150,142 200,138 Q250,134 300,138 Q350,142 400,138 Q450,134 500,138 Q550,142 600,138 Q650,134 700,138 Q750,142 800,138"
          stroke="rgba(255,255,255,0.7)"
          strokeWidth="0.9"
          fill="none"
        />
        <path
          d="M0,150 Q60,146 120,150 Q180,154 240,150 Q300,146 360,150 Q420,154 480,150 Q540,146 600,150 Q660,154 720,150 Q780,146 840,150"
          stroke="rgba(255,255,255,0.35)"
          strokeWidth="0.7"
          fill="none"
        />
      </g>
    </svg>
  );
}
