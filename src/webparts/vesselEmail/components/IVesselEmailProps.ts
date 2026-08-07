import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IVesselEmailProps {
  apiBaseUrl: string;
  apiToken: string;
  isDarkTheme: boolean;
  userDisplayName: string;
  userEmail: string;
  sessionId?: string;
  // True once the webpart's session-init flow has settled (success, no-session
  // stub mode, or failure). Data loading should gate on THIS, not on sessionId
  // being non-empty — a backend without a database legitimately never issues one.
  sessionInitialized?: boolean;

  // Graph / SharePoint context — required for folder creation and delta sync
  graphClient?: MSGraphClientV3;
  siteId?: string;
  driveId?: string;
}