import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'VesselEmailWebPartStrings';
import VesselEmail from './components/VesselEmail';
import { IVesselEmailProps } from './components/IVesselEmailProps';

export interface IVesselEmailWebPartProps {
  apiBaseUrl: string;
  apiToken: string;
}

export default class VesselEmailWebPart extends BaseClientSideWebPart<IVesselEmailWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _graphClient: MSGraphClientV3 | undefined;
  private _graphAccessToken: string | undefined;
  private _spAccessToken: string | undefined;  // SharePoint REST-scoped token for taxonomy writes
  private _siteId: string | undefined;
  private _driveId: string | undefined;
  private _userEmail: string = '';
  private _sessionId: string = '';
  private _sessionExpired: boolean = false;
  // True once the bypass-login call has SETTLED (success, no-session-stub-mode,
  // failure, or network error) — NOT the same as having a session id. The React
  // component uses this to know it's safe to load data even when the backend
  // is running without a database (where session_id is legitimately null).
  private _sessionInitialized: boolean = false;
  private _disposed: boolean = false;

  public render(): void {
    if (this._disposed || !this.domElement) {
      return;
    }
    const element: React.ReactElement<IVesselEmailProps> = React.createElement(
      VesselEmail,
      {
        apiBaseUrl: this._apiBaseUrl(),
        apiToken: this.properties.apiToken,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this._userEmail || this.context.pageContext.user.email,
        graphClient: this._graphClient,
        graphAccessToken: this._graphAccessToken,
        spAccessToken: this._spAccessToken,
        spHttpClient: this.context.spHttpClient,
        siteId: this._siteId,
        driveId: this._driveId,
        sessionId: this._sessionId,
        sessionExpired: this._sessionExpired,
        sessionInitialized: this._sessionInitialized,
        siteUrl: this.context.pageContext.site.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    }).then(() => this._initGraphContextAndSession());
  }

  private async _initGraphContextAndSession(): Promise<void> {
    // Step 1: resolve Graph context (email, siteId, driveId)
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      this._graphAccessToken = await tokenProvider.getToken('https://graph.microsoft.com');
      // Also fetch a SharePoint-scoped token; used by the backend for taxonomy (Vessel Name)
      // ValidateUpdateListItem calls which require the SP REST token audience.
      try {
        const spHost = this.context.pageContext.site.absoluteUrl
          ? new URL(this.context.pageContext.site.absoluteUrl).origin
          : '';
        if (spHost) {
          this._spAccessToken = await tokenProvider.getToken(spHost);
        }
      } catch (spTokErr) {
        console.warn('[VesselDMS] SP token fetch failed (non-fatal):', spTokErr);
      }
      const me = await this._graphClient.api('/me').select('mail,userPrincipalName').get();
      this._userEmail = me.mail || me.userPrincipalName || this.context.pageContext.user.email;

      const siteUrl = this.context.pageContext.site.absoluteUrl;
      const siteRes = await this._graphClient
        .api(`/sites/${new URL(siteUrl).hostname}:${new URL(siteUrl).pathname}`)
        .get();
      this._siteId = siteRes.id;
      const drivesRes = await this._graphClient.api(`/sites/${this._siteId}/drives`).get();
      const docLib = (drivesRes.value as any[]).find(
        (d: any) => d.driveType === 'documentLibrary' && d.name === 'Documents'
      ) || drivesRes.value[0];
      this._driveId = docLib?.id;
    } catch (e) {
      console.warn('[VesselDMS] Graph context init failed:', e);
      if (this._isExpiredGraphTokenError(e)) {
        this._sessionExpired = true;
        this._sessionInitialized = true;
        return;
      }
    }

    // Step 2: create a backend session BEFORE first render so _loadData has a valid session_id
    await this._ensureSession();
    // NOTE: Do NOT call this.render() here — BaseClientSideWebPart calls render()
    // automatically after onInit() resolves. Calling it manually causes a double
    // render where the first pass has sessionId='' and _loadData() bails out.
  }

  private _isExpiredGraphTokenError(error: unknown): boolean {
    const message = error instanceof Error ? error.message : String(error || '');
    const details = JSON.stringify(error || '');
    return /401|InvalidAuthenticationToken|Lifetime validation failed|token is expired|invalid authentication token/i.test(
      `${message} ${details}`
    );
  }

  private async _ensureSession(): Promise<void> {
    if (!this._userEmail) {
      // Fallback: use SharePoint page context email directly
      this._userEmail = this.context.pageContext.user.email || '';
    }
    if (!this._userEmail) {
      // No email could be resolved at all — still mark init as settled and
      // let the normal SPFx first render proceed.
      this._sessionInitialized = true;
      return;
    }
    const base = this._apiBaseUrl();
    try {
      let res: Response | null = null;
      try {
        res = await fetch(`${base}/api/auth/bypass-login`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email: this._userEmail,
            display_name: this.context.pageContext.user.displayName,
            tenant_id: this.context.pageContext.aadInfo?.tenantId?.toString() || '',
          }),
        });
      } catch (sslErr) {
        console.warn('[VesselDMS] Primary bypass-login failed (SSL/Network):', sslErr);
        if (base.includes('nk-dms-dev.sg-nissenkaiun.com')) {
          try {
            res = await fetch(`http://127.0.0.1:8000/api/auth/bypass-login`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                email: this._userEmail,
                display_name: this.context.pageContext.user.displayName,
                tenant_id: this.context.pageContext.aadInfo?.tenantId?.toString() || '',
              }),
            });
          } catch (localErr) {
            console.warn('[VesselDMS] Local fallback bypass-login failed:', localErr);
          }
        }
      }

      if (res && res.ok) {
        const data = await res.json();
        if (data.session_id) {
          this._sessionId = data.session_id;
        } else {
          console.warn('[VesselDMS] bypass-login succeeded without a session_id (stub/no-DB mode).');
        }
      } else if (res) {
        console.warn('[VesselDMS] bypass-login failed:', res.status, await res.text().catch(() => ''));
      }
    } catch (e) {
      console.warn('[VesselDMS] Session init failed:', e);
    } finally {


      // ALWAYS mark session init as settled — regardless of whether a
      // session_id was actually obtained. SPFx will render after onInit
      // resolves, so avoid manual render() here to prevent workbench/
      // property-pane lifecycle races.
      this._sessionInitialized = true;
    }
  }

  private _apiBaseUrl(): string {
    const configured = (this.properties.apiBaseUrl || '').replace(/\/$/, '');
    const hostedFromLocalDev = typeof window !== 'undefined' &&
      /(?:debugManifestsFile|localhost:4321)/i.test(window.location.href);
    if ((this.context.isServedFromLocalhost || hostedFromLocalDev) &&
        (!configured || configured === 'https://nk-dms-dev.sg-nissenkaiun.com')) {
      return 'http://127.0.0.1:8000';
    }
    return configured || 'https://nk-dms-dev.sg-nissenkaiun.com';
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    this._disposed = true;
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('apiBaseUrl', {
                  label: 'API Base URL'
                }),
                PropertyPaneTextField('apiToken', {
                  label: 'API Token'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}