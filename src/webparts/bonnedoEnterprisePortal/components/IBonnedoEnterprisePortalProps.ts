import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserPermissions } from '../models/PermissionModels';

export interface IBonnedoEnterprisePortalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  currentUserId: number;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  webPartContext: WebPartContext;
  userPermissions?: IUserPermissions;
  portalMode?: boolean;
  fullWidth?: boolean;
}
