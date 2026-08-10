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
  private _siteId: string | undefined;
  private _driveId: string | undefined;
  private _userEmail: string = '';
  private _sessionId: string = '';
  // True once the bypass-login call has SETTLED (success, no-session-stub-mode,
  // failure, or network error) — NOT the same as having a session id. The React
  // component uses this to know it's safe to load data even when the backend
  // is running without a database (where session_id is legitimately null).
  private _sessionInitialized: boolean = false;

  public render(): void {
    const element: React.ReactElement<IVesselEmailProps> = React.createElement(
      VesselEmail,
      {
        apiBaseUrl: this.properties.apiBaseUrl,
        apiToken: this.properties.apiToken,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this._userEmail || this.context.pageContext.user.email,
        graphClient: this._graphClient,
        siteId: this._siteId,
        driveId: this._driveId,
        sessionId: this._sessionId,
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
    }

    // Step 2: create a backend session BEFORE first render so _loadData has a valid session_id
    await this._ensureSession();
    // NOTE: Do NOT call this.render() here — BaseClientSideWebPart calls render()
    // automatically after onInit() resolves. Calling it manually causes a double
    // render where the first pass has sessionId='' and _loadData() bails out.
  }

  private async _ensureSession(): Promise<void> {
    if (!this._userEmail) {
      // Fallback: use SharePoint page context email directly
      this._userEmail = this.context.pageContext.user.email || '';
    }
    if (!this._userEmail) {
      // No email could be resolved at all — still mark init as settled and
      // render so the UI (and vessel list) isn't stuck waiting forever.
      this._sessionInitialized = true;
      this.render();
      return;
    }
    const base = (this.properties.apiBaseUrl || 'https://nk-dms-dev.sg-nissenkaiun.com').replace(/\/$/, '');
    try {
      const res = await fetch(`${base}/api/auth/bypass-login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email: this._userEmail,
          display_name: this.context.pageContext.user.displayName,
          tenant_id: this.context.pageContext.aadInfo?.tenantId?.toString() || '',
        }),
      });
      if (res.ok) {
        const data = await res.json();
        if (data.session_id) {
          this._sessionId = data.session_id;
        } else {
          // Backend is running in stub/no-DB mode (session_id intentionally null —
          // require_session() allows unauthenticated calls there). This is NOT a
          // failure: fall through and still mark init as done so the UI loads data.
          console.warn('[VesselDMS] bypass-login succeeded without a session_id (stub/no-DB mode).');
        }
      } else {
        console.warn('[VesselDMS] bypass-login failed:', res.status, await res.text().catch(() => ''));
        // Even without a session, the UI should still load — the backend allows
        // requests without a session token in stub/no-DB mode.
      }
    } catch (e) {
      console.warn('[VesselDMS] Session init failed:', e);
    } finally {
      // ALWAYS mark session init as settled and re-render — regardless of whether
      // a session_id was actually obtained. Previously this.render() was only
      // called inside the `if (data.session_id)` branch, so when the backend had
      // no database configured (session_id === null) this method returned
      // silently: the React component's props.sessionId stayed '' forever, its
      // _loadData() guard ("if (!this.props.sessionId) return;") never passed,
      // and the vessel list stayed blank with no error shown.
      this._sessionInitialized = true;
      this.render();
    }
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