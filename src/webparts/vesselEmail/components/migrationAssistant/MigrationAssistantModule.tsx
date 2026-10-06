// Entry point for the Migration Assistant module — the SharePoint AI
// Migration Assistant (originally a standalone Vite/React/Tailwind app at
// C:\llm\llm\SharePoint-AI-Migration-Assistant) ported to run as a tab
// inside this SPFx web part. Deliberately self-contained: everything it
// needs lives under this migrationAssistant/ folder, and the host app
// (VesselEmail.tsx) only ever touches this one file.
//
// The Migration Assistant backend is part of the main Vessel DMS backend
// (backend/app/migration_assistant/, routes under /api/migration-assistant/*)
// — it starts with the DMS API, no separate process. It reuses the DMS
// session for auth; the acting user's email is also sent for the audit trail.
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { MigrationApi } from './api';
import { MigrationScanTab } from './MigrationScanTab';
import { SiteToSiteTab } from './SiteToSiteTab';
import { TagExistingFilesTab } from './TagExistingFilesTab';
import { VesselExcelExportTab } from './VesselExcelExportTab';
import { tokens } from './styles';

export interface IMigrationAssistantModuleProps {
  /** Base URL of the main Vessel DMS API (the migration routes live on it).
   *  The web part's optional "Migration API Base URL" property overrides it. */
  apiBaseUrl: string;
  /** DMS session id (sent as Authorization / X-Session-ID). */
  sessionId?: string;
  actingEmail: string;
  isNight: boolean;
}

type Tab = 'migration' | 'site-to-site' | 'tag-existing' | 'excel-export';

const TABS: { id: Tab; label: string; icon: string }[] = [
  { id: 'migration', label: 'Migration Assistant', icon: 'Sparkle' },
  { id: 'site-to-site', label: 'Site-to-Site', icon: 'Switch' },
  { id: 'tag-existing', label: 'Tag Existing Files', icon: 'Tag' },
  { id: 'excel-export', label: 'Vessel Excel Export', icon: 'ExcelDocument' },
];

export function MigrationAssistantModule({ apiBaseUrl, sessionId, actingEmail, isNight }: IMigrationAssistantModuleProps): React.ReactElement {
  const t = tokens(isNight);
  const [tab, setTab] = React.useState<Tab>('migration');
  const [health, setHealth] = React.useState<'checking' | 'ok' | 'not_configured' | 'unreachable'>('checking');

  const api = React.useMemo(() => new MigrationApi(apiBaseUrl, actingEmail || 'unknown', sessionId || ''), [apiBaseUrl, actingEmail, sessionId]);

  React.useEffect(() => {
    let cancelled = false;
    setHealth('checking');
    // "configured: false" means the backend is up but has no Migration
    // Assistant settings (MIGRATION_* in backend/.env) — say so instead of green.
    api.health()
      .then((h) => !cancelled && setHealth(h.configured ? 'ok' : 'not_configured'))
      .catch(() => !cancelled && setHealth('unreachable'));
    return () => { cancelled = true; };
  }, [api]);

  if (!apiBaseUrl) {
    return (
      <div style={{ padding: 40, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>
        <Icon iconName="Warning" style={{ fontSize: 24, marginBottom: 10, display: 'block', margin: '0 auto 10px' }} />
        The Migration Assistant runs on the main DMS backend. Set "API Base URL" in this web part's property pane
        (the pencil / edit-web-part panel) to your DMS API (default local dev: <code>http://127.0.0.1:8000</code>).
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', padding: '0 20px 20px' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 4, borderBottom: `1px solid ${t.border}`, paddingTop: 8 }}>
        {TABS.map((tb) => (
          <button
            key={tb.id}
            onClick={() => setTab(tb.id)}
            style={{
              display: 'flex', alignItems: 'center', gap: 6, border: 'none', background: 'transparent',
              borderBottom: tab === tb.id ? '2px solid var(--clay-accent)' : '2px solid transparent',
              color: tab === tb.id ? 'var(--clay-accent)' : t.textMuted, fontWeight: tab === tb.id ? 700 : 500,
              fontSize: 13, padding: '10px 12px', cursor: 'pointer',
            }}
          >
            <Icon iconName={tb.icon} style={{ fontSize: 13 }} />{tb.label}
          </button>
        ))}
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: t.textSubtle, paddingRight: 4 }}>
          <span style={{ width: 7, height: 7, borderRadius: 999, background: health === 'ok' ? '#16a34a' : health === 'unreachable' ? '#dc2626' : '#d97706' }} />
          {health === 'ok' ? 'Migration Assistant connected'
            : health === 'checking' ? 'Checking Migration Assistant…'
            : health === 'not_configured' ? 'Migration Assistant not configured on the server (set MIGRATION_SITE_HOSTNAME / MIGRATION_SITE_PATH in backend .env)'
            : `Migration Assistant unreachable at ${apiBaseUrl}`}
        </div>
      </div>
      <div style={{ flex: 1, overflow: 'hidden', paddingTop: 14, display: 'flex', flexDirection: 'column' }}>
        {tab === 'migration' && <MigrationScanTab api={api} isNight={isNight} actingEmail={actingEmail} />}
        {tab === 'site-to-site' && <SiteToSiteTab api={api} isNight={isNight} />}
        {tab === 'tag-existing' && <TagExistingFilesTab api={api} isNight={isNight} />}
        {tab === 'excel-export' && <VesselExcelExportTab api={api} isNight={isNight} />}
      </div>
    </div>
  );
}
