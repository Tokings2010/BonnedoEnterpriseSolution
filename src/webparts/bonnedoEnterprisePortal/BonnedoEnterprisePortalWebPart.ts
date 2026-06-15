import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BonnedoEnterprisePortalWebPartStrings';
import BonnedoEnterprisePortal from './components/BonnedoEnterprisePortal';
import { IBonnedoEnterprisePortalProps } from './components/IBonnedoEnterprisePortalProps';
import { PermissionService } from './services/PermissionService';
import { IUserPermissions } from './models/PermissionModels';

export interface IBonnedoEnterprisePortalWebPartProps {
  description: string;
  portalMode: boolean;
  fullWidth: boolean;
}

export default class BonnedoEnterprisePortalWebPart extends BaseClientSideWebPart<IBonnedoEnterprisePortalWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _userPermissions: IUserPermissions | undefined = undefined;
  private _permissionService: PermissionService | undefined = undefined;

  public render(): void {
    const element: React.ReactElement<IBonnedoEnterprisePortalProps> = React.createElement(
      BonnedoEnterprisePortal,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        currentUserId: parseInt((this.context.pageContext.user as any).id?.toString() || '0', 10),
        spHttpClient: this.context.spHttpClient,
        pageContext: this.context.pageContext,
        webPartContext: this.context,
        userPermissions: this._userPermissions,
        portalMode: this.properties.portalMode,
        fullWidth: this.properties.fullWidth
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Initialize PermissionService
    this._permissionService = new PermissionService(this.context.spHttpClient, this.context.pageContext);

    // Get user permissions
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      return this._loadUserPermissions();
    });
  }

  private async _loadUserPermissions(): Promise<void> {
    try {
      if (this._permissionService) {
        this._userPermissions = await this._permissionService.getUserPermissions();
        console.log('User permissions loaded:', this._userPermissions);
      }
    } catch (error) {
      console.error('Error loading user permissions:', error);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('portalMode', {
                  label: 'Portal Mode',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('fullWidth', {
                  label: 'Full Width',
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
