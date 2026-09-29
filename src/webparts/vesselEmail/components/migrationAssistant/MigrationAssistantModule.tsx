// Entry point for the Migration Assistant module — the SharePoint AI
// Migration Assistant (originally a standalone Vite/React/Tailwind app at
// C:\llm\llm\SharePoint-AI-Migration-Assistant) ported to run as a tab
// inside this SPFx web part. Deliberately self-contained: everything it
// needs lives under this migrationAssistant/ folder, and the host app
// (VesselEmail.tsx) only ever touches this one file.
//
// It talks to the migration backend's OWN FastAPI service — a separate
// process from this SPFx app's main Vessel DMS backend, run per this
// module's README.md. It shares no code, database, or auth session with
// the main app; the only thing passed in is the acting user's email (for
// the migration backend's own audit trail) and a base URL to reach it at.
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { MigrationApi } from './api';
import { MigrationScanTab } from './MigrationScanTab';
import { SiteToSiteTab } from './SiteToSiteTab';
import { TagExistingFilesTab } from './TagExistingFilesTab';
import { VesselExcelExportTab } from './VesselExcelExportTab';
import { tokens } from './styles';

export interface IMigrationAssistantModuleProps {
  /** Base URL of the standalone migration backend, e.g. http://localhost:8020
   *  or https://your-migration-host. Configured via the web part's property
   *  pane ("Migration API Base URL"); falls back to localhost:8020 when
   *  served from the local SPFx workbench. */
  apiBaseUrl: string;
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

export function MigrationAssistantModule({ apiBaseUrl, actingEmail, isNight }: IMigrationAssistantModuleProps): React.ReactElement {
  const t = tokens(isNight);
  const [tab, setTab] = React.useState<Tab>('migration');
  const [health, setHealth] = React.useState<'checking' | 'ok' | 'unreachable'>('checking');

  const api = React.useMemo(() => new MigrationApi(apiBaseUrl, actingEmail || 'unknown'), [apiBaseUrl, actingEmail]);

  React.useEffect(() => {
    let cancelled = false;
    setHealth('checking');
    api.health().then(() => !cancelled && setHealth('ok')).catch(() => !cancelled && setHealth('unreachable'));
    return () => { cancelled = true; };
  }, [api]);

  if (!apiBaseUrl) {
    return (
      <div style={{ padding: 40, textAlign: 'center', color: t.textMuted, fontSize: 13 }}>
        <Icon iconName="Warning" style={{ fontSize: 24, marginBottom: 10, display: 'block', margin: '0 auto 10px' }} />
        The Migration Assistant needs its backend's URL. Set "Migration API Base URL" in this web part's property pane
        (the pencil / edit-web-part panel), pointing at the migration backend documented in this module's README.md
        (default local dev: <code>http://localhost:8020</code>).
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
              borderBottom: tab === tb.id ? '2px solid #DD9159' : '2px solid transparent',
              color: tab === tb.id ? '#DD9159' : t.textMuted, fontWeight: tab === tb.id ? 700 : 500,
              fontSize: 13, padding: '10px 12px', cursor: 'pointer',
            }}
          >
            <Icon iconName={tb.icon} style={{ fontSize: 13 }} />{tb.label}
          </button>
        ))}
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: t.textSubtle, paddingRight: 4 }}>
          <span style={{ width: 7, height: 7, borderRadius: 999, background: health === 'ok' ? '#16a34a' : health === 'checking' ? '#d97706' : '#dc2626' }} />
          {health === 'ok' ? 'Migration backend connected' : health === 'checking' ? 'Checking migration backend…' : `Migration backend unreachable at ${apiBaseUrl}`}
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
